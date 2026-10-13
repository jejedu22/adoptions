<?php

namespace App\Tests\Import;

use App\Entity\Decision;
use App\Entity\Demande;
use App\Entity\Dossier;
use App\Entity\EnfantDemande;
use App\Entity\HistoriqueIntervenant;
use App\Entity\Intervenant;
use App\Entity\Placement;
use App\Entity\Surveillance;
use Doctrine\DBAL\DriverManager;
use Doctrine\DBAL\Tools\DsnParser;
use Doctrine\ORM\EntityManagerInterface;
use Doctrine\ORM\Tools\SchemaTool;
use Symfony\Bundle\FrameworkBundle\Console\Application;
use Symfony\Bundle\FrameworkBundle\Test\KernelTestCase;
use Symfony\Component\Console\Tester\CommandTester;

class ImportLegacyCommandTest extends KernelTestCase
{
    private string $legacyFile;

    protected function setUp(): void
    {
        $this->legacyFile = sys_get_temp_dir().'/legacy_'.uniqid().'.sqlite';
        $db = DriverManager::getConnection((new DsnParser(['sqlite' => 'pdo_sqlite']))->parse('sqlite:///'.$this->legacyFile));
        foreach ([
            'CREATE TABLE TYPE (TYPE_CODE TEXT, TYPE_LIBELLE TEXT)',
            "INSERT INTO TYPE VALUES ('I','Initiale')",
            'CREATE TABLE TYRE (TYRE_NUM INT, TYRE_NOM TEXT)',
            "INSERT INTO TYRE VALUES (1,'Tribunal administratif'),(3,'Conseil d''État')",
            'CREATE TABLE TRSO (TRSO_NUM INT, TRSO_NOM TEXT, TRSO_ACTIF INT)',
            "INSERT INTO TRSO VALUES (176,'COLLIN',-1),(177,'MARTIN',-1)",
            'CREATE TABLE DOSS (DOSS_NUM INT, DOSS_VILLE TEXT, DOSS_REINS INT)',
            "INSERT INTO DOSS VALUES (9907903,'Nancy',0)",
            'CREATE TABLE INDI (INDI_NUM INT, DOSS_NUM INT, INDI_CDFR TEXT, INDI_NOM TEXT, INDI_PREN TEXT, INDI_DNAIS TEXT)',
            "INSERT INTO INDI VALUES (1,9907903,'DE','DUPONT','Marie','1970-05-02 00:00:00'),(2,424242,'DE','ORPHELIN','X',NULL)",
            'CREATE TABLE DEMA (DEMA_NUM INT, DOSS_NUM INT, TYPE_CODE TEXT, DEMA_DATE TEXT, DEMA_CFRAT INT)',
            "INSERT INTO DEMA VALUES (9907903,9907903,'I','2003-04-04 00:00:00',-1)",
            'CREATE TABLE u02_HistoCommissions (DEMA_NUM INT, u02_dtmHistoCommissionReunion TEXT, u02_dtmHistoCommissionSignature TEXT, u02_intHistoCommissionDecision INT)',
            "INSERT INTO u02_HistoCommissions VALUES (9907903,'2003-04-04 00:00:00','2003-05-01 00:00:00',1)",
            'CREATE TABLE RECO (TYRE_NUM INT, DEMA_NUM INT, u02_dtmHistoCommissionReunion TEXT, RECO_DDREC TEXT, RECO_DECI TEXT, RECO_dtmAgrementDelivrance TEXT)',
            "INSERT INTO RECO VALUES (1,9907903,'2003-04-04 00:00:00','2003-09-10 00:00:00','A','2004-01-27 00:00:00')",
            'CREATE TABLE u01_HistoIntervenants (DEMA_NUM INT, TRSO_NUM INT, u01_intHistoIntervenantCadre INT, u01_booHistoIntervenantPsychologue INT, u01_dtmHistoIntervenantDemande TEXT)',
            "INSERT INTO u01_HistoIntervenants VALUES (9907903,176,1,0,'2003-04-10 00:00:00'),(9907903,177,1,-1,'2003-04-10 00:00:00')",
            'CREATE TABLE ENDE (ENDE_NUM INT, ENDE_SEXE TEXT, ENDE_HAND TEXT, ENDE_TRAGE TEXT)',
            "INSERT INTO ENDE VALUES (5108,'F','O','00-03')",
            'CREATE TABLE FADE (DEMA_NUM INT, ENDE_NUM INT)',
            'INSERT INTO FADE VALUES (9907903,5108)',
            'CREATE TABLE PLAC (PLAC_NUM INT, DEMA_NUM INT, PLAC_DATE TEXT, PLAC_PUPILL INT)',
            "INSERT INTO PLAC VALUES (541,9907903,'2004-06-01 00:00:00',-1)",
            'CREATE TABLE ENPL (PLAC_NUM INT, ENPL_NOM TEXT, ENPL_PREN TEXT, ENPL_SEXE TEXT)',
            "INSERT INTO ENPL VALUES (541,'Petit','Léa','F')",
            'CREATE TABLE SUPL (PLAC_NUM INT, TRSO_NUMAS INT, SUPL_DDSUR TEXT)',
            "INSERT INTO SUPL VALUES (541,176,'2004-10-03 00:00:00')",
        ] as $sql) {
            $db->executeStatement($sql);
        }
    }

    protected function tearDown(): void
    {
        @unlink($this->legacyFile);
        parent::tearDown();
    }

    private function lancer(array $options): CommandTester
    {
        $kernel = self::bootKernel();
        $em = static::getContainer()->get(EntityManagerInterface::class);
        $tool = new SchemaTool($em);
        $tool->dropSchema($em->getMetadataFactory()->getAllMetadata());
        $tool->createSchema($em->getMetadataFactory()->getAllMetadata());

        return $this->execute($kernel, $options);
    }

    private function execute($kernel, array $options): CommandTester
    {
        $tester = new CommandTester((new Application($kernel))->find('app:import-legacy'));
        $tester->execute(['--source' => 'sqlite:///'.$this->legacyFile] + $options);

        return $tester;
    }

    public function testImportComplet(): void
    {
        $tester = $this->lancer([]);
        self::assertSame(0, $tester->getStatusCode(), $tester->getDisplay());
        $em = static::getContainer()->get(EntityManagerInterface::class);
        $em->clear();

        $dossier = $em->getRepository(Dossier::class)->findOneBy(['numero' => '9907903']);
        self::assertSame('Nancy', $dossier->getVille());
        self::assertCount(1, $dossier->getIndividus(), 'L\'individu sans dossier est ignoré');
        self::assertSame('1970-05-02', $dossier->getIndividus()->first()->getDateNaissance()->format('Y-m-d'));

        $demande = $em->getRepository(Demande::class)->findOneBy(['numero' => 9907903]);
        self::assertTrue($demande->isFratrie());
        self::assertSame('Initiale', (string) $demande->getType());
        self::assertCount(2, $demande->getDecisions());
        // refus de la commission puis accord sur recours contentieux -> agrément recalculé
        self::assertSame('2004-01-27', $demande->getDateAgrement()?->format('Y-m-d'));
        self::assertSame(1, $demande->getProvenanceAgrement());

        $types = [];
        foreach ($em->getRepository(Intervenant::class)->findAll() as $i) {
            $types[$i->getNom()] = $i->getType();
        }
        self::assertSame(['COLLIN' => 'assistante_sociale', 'MARTIN' => 'psychologue'], $types);

        $enfants = $em->getRepository(EnfantDemande::class)->findAll();
        self::assertCount(1, $enfants);
        self::assertSame(3, $enfants[0]->getSexe());
        self::assertTrue($enfants[0]->isHandicap());
        self::assertSame(3, $enfants[0]->getAgeMax());

        $placement = $em->getRepository(Placement::class)->findOneBy([]);
        self::assertSame('Petit', $placement->getNomEnfant());
        self::assertTrue($placement->isPupille());
        self::assertCount(1, $placement->getSurveillances());
        self::assertCount(2, $em->getRepository(HistoriqueIntervenant::class)->findAll());
        self::assertTrue($em->getRepository(Decision::class)->findOneBy(['provenance' => 3]) === null);
        self::assertStringContainsString('absente', $tester->getDisplay(), 'les tables manquantes sont signalées');
    }

    public function testRelanceSansDoublon(): void
    {
        $kernel = self::bootKernel();
        $this->lancer([]);
        $this->execute($kernel, []);
        $em = static::getContainer()->get(EntityManagerInterface::class);
        self::assertSame(1, $em->getRepository(Dossier::class)->count([]));
        self::assertSame(2, $em->getRepository(Decision::class)->count([]));
        self::assertSame(2, $em->getRepository(HistoriqueIntervenant::class)->count([]));
        self::assertSame(1, $em->getRepository(Surveillance::class)->count([]));
    }

    public function testDryRunEtCheckNEcriventRien(): void
    {
        $tester = $this->lancer(['--dry-run' => true]);
        self::assertSame(0, $tester->getStatusCode(), $tester->getDisplay());
        $em = static::getContainer()->get(EntityManagerInterface::class);
        self::assertSame(0, $em->getRepository(Dossier::class)->count([]));

        $tester = $this->execute(self::bootKernel(), ['--check' => true]);
        self::assertStringContainsString('Vérification seule', $tester->getDisplay());
        self::assertSame(0, $em->getRepository(Dossier::class)->count([]));
    }

    public function testStrictEchoueSiColonnesAbsentes(): void
    {
        $tester = $this->lancer(['--strict' => true]);
        self::assertSame(1, $tester->getStatusCode());
    }
}

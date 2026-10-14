<?php

namespace App\Tests\Controller;

use App\Entity\Dossier;
use Doctrine\ORM\EntityManagerInterface;
use Doctrine\ORM\Tools\SchemaTool;
use Symfony\Bundle\FrameworkBundle\Test\WebTestCase;
use Symfony\Component\Security\Core\User\InMemoryUser;

class SmokeTest extends WebTestCase
{
    public function testParcoursPrincipal(): void
    {
        $client = static::createClient();
        $em = static::getContainer()->get(EntityManagerInterface::class);
        $tool = new SchemaTool($em);
        $tool->dropSchema($em->getMetadataFactory()->getAllMetadata());
        $tool->createSchema($em->getMetadataFactory()->getAllMetadata());

        $client->request('GET', '/');
        self::assertResponseStatusCodeSame(401);

        $client->loginUser(new InMemoryUser('admin', $_SERVER['APP_ADMIN_PASSWORD_HASH'], ['ROLE_USER']));
        $client->request('GET', '/');
        self::assertResponseIsSuccessful();

        foreach (['/dossiers', '/dossiers/nouveau', '/a-traiter', '/commissions', '/placements', '/recherche?q=du', '/administration', '/gestion/professions', '/gestion/professions/nouveau', '/gestion/parametres'] as $url) {
            $client->request('GET', $url);
            self::assertResponseIsSuccessful($url);
        }

        $crawler = $client->request('GET', '/dossiers/nouveau');
        $form = $crawler->selectButton('Créer le dossier ✓')->form(['nouveau_dossier[numero]' => 'D-001', 'nouveau_dossier[nom]' => 'Dupont', 'nouveau_dossier[prenom]' => 'Marie', 'nouveau_dossier[ville]' => 'Nancy', 'nouveau_dossier[dateDemande]' => '2024-02-01']);
        $client->submit($form);
        self::assertResponseRedirects();
        $dossier = $em->getRepository(Dossier::class)->findOneBy(['numero' => 'D-001']);
        self::assertNotNull($dossier);

        $client->request('GET', '/dossiers/'.$dossier->getId());
        self::assertResponseIsSuccessful();
        self::assertSelectorTextContains('h1', 'Dupont Marie');
        self::assertSame(1, $dossier->getDemandes()->count());

        $client->request('GET', '/recherche?q=Dupont');
        self::assertSelectorTextContains('body', 'Dupont Marie');
        $client->request('GET', '/dossiers?statut=instruction');
        self::assertSelectorTextContains('body', 'D-001');

        $demande = $em->getRepository(\App\Entity\Demande::class)->findOneBy([]);
        $client->request('GET', '/demandes/'.$demande->getId());
        self::assertResponseIsSuccessful();
        self::assertSelectorTextContains('h1', 'Demande n° 1');

        $crawler = $client->request('GET', '/gestion/decisions/nouveau?parent='.$demande->getId());
        self::assertResponseIsSuccessful();
        $form = $crawler->selectButton('Valider')->form([
            'entity[dateCommission]' => '2024-03-10',
            'entity[commissionInitiale]' => '2024-03-10',
            'entity[deliberation]' => '3',
            'entity[provenance]' => '-2',
        ]);
        $client->submit($form);
        self::assertResponseRedirects();
        $em->clear();
        $demande = $em->getRepository(\App\Entity\Demande::class)->findOneBy([]);
        self::assertSame('2024-03-10', $demande->getDateAgrement()?->format('Y-m-d'));
    }

    public function testEnTetesDuProxyInverse(): void
    {
        $client = static::createClient();
        $client->loginUser(new InMemoryUser('admin', $_SERVER['APP_ADMIN_PASSWORD_HASH'], ['ROLE_USER']));
        $client->request('GET', '/', server: ['REMOTE_ADDR' => '127.0.0.1', 'HTTP_X_FORWARDED_PROTO' => 'https', 'HTTP_X_FORWARDED_HOST' => 'adoptions.example.org', 'HTTP_X_FORWARDED_PORT' => '443']);
        self::assertTrue($client->getRequest()->isSecure());
        self::assertSame('adoptions.example.org', $client->getRequest()->getHost());
    }

    public function testRapportEnRetardEtActionRapide(): void
    {
        $client = static::createClient();
        $em = static::getContainer()->get(EntityManagerInterface::class);
        $tool = new SchemaTool($em);
        $tool->dropSchema($em->getMetadataFactory()->getAllMetadata());
        $tool->createSchema($em->getMetadataFactory()->getAllMetadata());

        $dossier = (new Dossier())->setNumero('R-1');
        $dossier->addIndividu((new \App\Entity\Individu())->setType('DE')->setNom('Roux')->setPrenom('Lea'));
        $demande = (new \App\Entity\Demande())->setNumero(1)->setDateDemande(new \DateTimeImmutable('today -200 days'));
        $dossier->addDemande($demande);
        $psy = (new \App\Entity\Intervenant())->setNom('MARTIN')->setType('psychologue');
        $h = (new \App\Entity\HistoriqueIntervenant())->setIntervenant($psy)->setCadre(1)->setDatePrevue(new \DateTimeImmutable('today -10 days'));
        $demande->addHistorique($h);
        $em->persist($dossier);
        $em->persist($psy);
        $em->flush();

        $client->loginUser(new InMemoryUser('admin', $_SERVER['APP_ADMIN_PASSWORD_HASH'], ['ROLE_USER']));
        $crawler = $client->request('GET', '/a-traiter');
        self::assertSelectorTextContains('body', 'rapport psychologue attendu');
        self::assertSelectorTextContains('body', 'En retard · 10 j');
        $client->request('GET', '/');
        self::assertSelectorTextContains('aside', '1');

        $client->submit($crawler->selectButton('Reçu aujourd\'hui')->form());
        self::assertResponseRedirects();
        $em->clear();
        self::assertNotNull($em->getRepository(\App\Entity\HistoriqueIntervenant::class)->findOneBy([])->getDateRecu());
        $client->request('GET', '/a-traiter');
        self::assertSelectorTextContains('body', 'Aucune action');
    }

    public function testFiltresDesCommissions(): void
    {
        $client = static::createClient();
        $em = static::getContainer()->get(EntityManagerInterface::class);
        $tool = new SchemaTool($em);
        $tool->dropSchema($em->getMetadataFactory()->getAllMetadata());
        $tool->createSchema($em->getMetadataFactory()->getAllMetadata());

        foreach ([['A-1', 'Alpha', 3, '2024-03-10', -2], ['B-2', 'Bravo', 1, '2025-06-01', -2], ['C-3', 'Charlie', 0, '2025-09-01', -1]] as [$num, $nom, $delib, $date, $prov]) {
            $dossier = (new Dossier())->setNumero($num);
            $dossier->addIndividu((new \App\Entity\Individu())->setType('DE')->setNom($nom)->setPrenom('X'));
            $demande = (new \App\Entity\Demande())->setNumero(1);
            $dossier->addDemande($demande);
            $demande->addDecision((new \App\Entity\Decision())->setProvenance($prov)->setDeliberation($delib)->setDateCommission(new \DateTimeImmutable($date)));
            $em->persist($dossier);
        }
        $em->flush();

        $client->loginUser(new InMemoryUser('admin', $_SERVER['APP_ADMIN_PASSWORD_HASH'], ['ROLE_USER']));
        $noms = static function ($crawler) { return $crawler->filter('tbody td a.strong')->each(fn ($n) => $n->text()); };

        self::assertCount(3, $noms($client->request('GET', '/commissions')));
        self::assertSame(['Alpha X'], $noms($client->request('GET', '/commissions?dossier=alph')));
        self::assertSame(['Bravo X'], $noms($client->request('GET', '/commissions?decision=1')));
        self::assertSame(['Charlie X'], $noms($client->request('GET', '/commissions?decision=0')));
        self::assertSame(['Charlie X'], $noms($client->request('GET', '/commissions?origine=administratif')));
        self::assertSame(['Charlie X', 'Bravo X'], $noms($client->request('GET', '/commissions?commission_du=2025-01-01')));
        self::assertSame(['Alpha X'], $noms($client->request('GET', '/commissions?commission_au=2024-12-31')));
        self::assertSame(['Bravo X'], $noms($client->request('GET', '/commissions?commission_du=2025-01-01&commission_au=2025-07-01&origine=initiale')));
        self::assertSame([], $noms($client->request('GET', '/commissions?dossier=zzz')));
        self::assertResponseIsSuccessful();
    }
}

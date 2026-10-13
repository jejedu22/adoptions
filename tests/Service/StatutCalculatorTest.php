<?php

namespace App\Tests\Service;

use App\Entity\Decision;
use App\Entity\Demande;
use App\Entity\HistoriqueIntervenant;
use App\Entity\Intervenant;
use App\Service\Statut;
use App\Service\StatutCalculator;
use PHPUnit\Framework\TestCase;

class StatutCalculatorTest extends TestCase
{
    private function d(string $s): \DateTimeImmutable
    {
        return new \DateTimeImmutable($s);
    }

    public function testDemandeEnInstructionAvecRapportAttendu(): void
    {
        $demande = (new Demande())->setDateDemande($this->d('2025-01-10'))->setDateReceptionDossier($this->d('2025-01-20'));
        $psy = (new Intervenant())->setNom('MARTIN');
        $demande->getHistoriques()->add((new HistoriqueIntervenant())->setDemande($demande)->setIntervenant($psy)->setCadre(1)->setDatePrevue($this->d('2025-03-01')));
        $demande->getHistoriques()->add((new HistoriqueIntervenant())->setDemande($demande)->setIntervenant($psy)->setCadre(1)->setDateRecu($this->d('2025-02-01')));

        $s = (new StatutCalculator())->calculer($demande, $this->d('2025-04-01'));
        self::assertSame('En instruction', $s->libelle);
        self::assertSame(Statut::WARN, $s->ton);
        self::assertSame('Rapport attendu : MARTIN', $s->prochaineEtape);
        self::assertSame(['done', 'done', 'cur', 'todo', 'todo'], array_column($s->etapes, 'etat'));
        self::assertSame('1 sur 2', $s->etapes[2]['detail']);
    }

    public function testAgreeEtReactualisation(): void
    {
        $demande = (new Demande())->setDateDemande($this->d('2022-01-01'))->setDateAgrement($this->d('2022-06-01'))->setAActualiser(true);
        $calc = new StatutCalculator();
        self::assertSame('Agréé', $calc->calculer($demande, $this->d('2023-12-01'))->libelle);
        self::assertSame('À réactualiser', $calc->calculer($demande, $this->d('2024-03-01'))->libelle);
    }

    public function testRefusAvecRecours(): void
    {
        $demande = (new Demande())->setDateDemande($this->d('2022-01-01'));
        $demande->getDecisions()->add((new Decision())->setDemande($demande)->setProvenance(-2)->setDeliberation(1)->setCommissionInitiale($this->d('2022-05-01')));
        $demande->getDecisions()->add((new Decision())->setDemande($demande)->setProvenance(-1)->setDeliberation(0)->setDateDepot($this->d('2022-06-01')));
        self::assertSame('Refus · recours', (new StatutCalculator())->calculer($demande)->libelle);
    }

    public function testSortieEtSuspension(): void
    {
        $calc = new StatutCalculator();
        self::assertSame('Sorti', $calc->calculer((new Demande())->setDateSortie($this->d('2024-01-01')))->libelle);
        self::assertSame('Suspendu', $calc->calculer((new Demande())->setDateSuspension($this->d('2024-01-01')))->libelle);
    }
}

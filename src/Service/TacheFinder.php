<?php

namespace App\Service;

use App\Entity\Decision;
use App\Entity\Demande;
use App\Entity\Enum\Deliberation;
use App\Entity\HistoriqueIntervenant;
use App\Entity\Surveillance;
use Doctrine\ORM\EntityManagerInterface;

/** Rassemble toutes les échéances à traiter, sans avoir à ouvrir chaque dossier. */
final class TacheFinder
{
    public function __construct(private readonly EntityManagerInterface $em)
    {
    }

    /**
     * @param int $horizonJours échéances à venir prises en compte (en plus de tout ce qui est en retard)
     *
     * @return list<Tache> triées : les plus en retard d'abord
     */
    public function trouver(int $horizonJours = 30, ?\DateTimeImmutable $aujourdhui = null): array
    {
        $aujourdhui ??= new \DateTimeImmutable('today');
        $limite = $aujourdhui->modify("+$horizonJours days");
        $taches = [];
        $jours = static fn (\DateTimeImmutable $d): int => (int) $aujourdhui->diff($d)->format('%r%a') * -1;

        // Rapports et interventions attendus
        $historiques = $this->em->createQuery(
            'SELECT h, d, i FROM App\Entity\HistoriqueIntervenant h JOIN h.demande d JOIN h.intervenant i
             WHERE h.dateRecu IS NULL AND h.datePrevue IS NOT NULL AND h.datePrevue <= :limite AND d.dateSortie IS NULL'
        )->setParameter('limite', $limite)->getResult();
        /** @var HistoriqueIntervenant $h */
        foreach ($historiques as $h) {
            $role = 'psychologue' === $h->getIntervenant()->getType() ? 'psychologue' : 'travailleur social';
            $taches[] = new Tache('rapport', 'Rapport '.$role.' attendu', 'Intervenant : '.$h->getIntervenant()->getNom(), $h->getDemande(), $h->getDatePrevue(), $jours($h->getDatePrevue()), $h->getId(), $h->getIntervenant()->getNom());
        }

        // Décisions prévues mais pas encore saisies
        $decisions = $this->em->createQuery(
            'SELECT x, d FROM App\Entity\Decision x JOIN x.demande d
             WHERE x.deliberation = :ind AND x.dateDecisionPrevue IS NOT NULL AND x.dateDecisionPrevue <= :limite AND d.dateSortie IS NULL'
        )->setParameter('ind', Deliberation::Indetermine->value)->setParameter('limite', $limite)->getResult();
        /** @var Decision $x */
        foreach ($decisions as $x) {
            $taches[] = new Tache('decision', 'Décision à saisir', 'Commission / recours prévu(e)', $x->getDemande(), $x->getDateDecisionPrevue(), $jours($x->getDateDecisionPrevue()));
        }

        // Surveillances de placement à réaliser
        $surveillances = $this->em->createQuery(
            'SELECT s, p, d FROM App\Entity\Surveillance s JOIN s.placement p JOIN p.demande d
             WHERE s.dateRealisee IS NULL AND s.datePrevue IS NOT NULL AND s.datePrevue <= :limite'
        )->setParameter('limite', $limite)->getResult();
        /** @var Surveillance $s */
        foreach ($surveillances as $s) {
            $taches[] = new Tache('surveillance', 'Surveillance du placement', trim('Enfant : '.$s->getPlacement()->getNomEnfant().' '.$s->getPlacement()->getPrenomEnfant()), $s->getPlacement()->getDemande(), $s->getDatePrevue(), $jours($s->getDatePrevue()), null, $s->getIntervenant()?->getNom());
        }

        // Agréments à réactualiser
        $seuil = $aujourdhui->modify('-'.StatutCalculator::MOIS_REACTUALISATION.' months');
        $demandes = $this->em->createQuery(
            'SELECT d FROM App\Entity\Demande d WHERE d.dateAgrement IS NOT NULL AND d.dateAgrement < :seuil AND d.aActualiser = true AND d.dateSortie IS NULL AND d.dateRetrait IS NULL'
        )->setParameter('seuil', $seuil)->getResult();
        /** @var Demande $d */
        foreach ($demandes as $d) {
            $echeance = $d->getDateAgrement()->modify('+'.StatutCalculator::MOIS_REACTUALISATION.' months');
            $taches[] = new Tache('reactualisation', 'Agrément à réactualiser', 'Délivré le '.$d->getDateAgrement()->format('d/m/Y'), $d, $echeance, $jours($echeance));
        }

        usort($taches, static fn (Tache $a, Tache $b) => $b->retard <=> $a->retard);

        return $taches;
    }

    public function nombreEnRetard(): int
    {
        return count(array_filter($this->trouver(0), static fn (Tache $t) => $t->enRetard()));
    }
}

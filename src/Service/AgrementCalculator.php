<?php

namespace App\Service;

use App\Entity\Decision;
use App\Entity\Demande;
use App\Entity\Enum\Deliberation;

/**
 * Détermine, à partir de l'ensemble des décisions (commission initiale, recours administratif,
 * recours contentieux) d'une demande, la date et la provenance de l'agrément, ainsi que la date
 * de retrait éventuelle.
 *
 * Trois types d'enregistrements coexistent :
 *  - commission : dateCommission toujours renseignée, dateDepot nulle ;
 *  - recours administratif : dateDepot renseignée, dateCommission nulle si la décision est prise
 *    sans commission (la dateSignature fait alors foi) ;
 *  - recours contentieux : dateCommission nulle, dateSignature = date de l'agrément.
 */
final class AgrementCalculator
{
    public const PROVENANCE_DEMANDE_INITIALE = -2;
    public const PROVENANCE_RECOURS_ADMINISTRATIF = -1;

    /** @param iterable<Decision> $decisions décisions triées par date de commission puis identifiant */
    public function calculer(iterable $decisions): AgrementResult
    {
        $premiereCommission = null;
        $agrementDate = null;
        $agrementProvenance = null;
        $retraitInitial = null;
        $retrait = null;

        foreach ($decisions as $decision) {
            $deliberation = Deliberation::tryFrom($decision->getDeliberation()) ?? Deliberation::Indetermine;
            $conseilEtat = $decision->getTypeRecours()?->isConseilEtat() ?? false;

            if (Decision::isInitiale($decision) && null === $premiereCommission && Deliberation::Indetermine !== $deliberation) {
                $premiereCommission = $decision->getCommissionInitiale();
            }

            $commission = $decision->getCommissionInitiale();
            $memeCommission = null !== $commission && null !== $premiereCommission && $commission == $premiereCommission;

            if ($memeCommission) {
                switch ($deliberation) {
                    case Deliberation::Indetermine:
                        if (!$conseilEtat) {
                            $agrementDate = null;
                            $agrementProvenance = null;
                        }
                        break;
                    case Deliberation::Refus:
                        $agrementDate = null;
                        $agrementProvenance = null;
                        break;
                    case Deliberation::Accord:
                        if (!$conseilEtat || null === $agrementDate) {
                            $agrementDate = $decision->getDateCommission() ?? $decision->getDateSignature();
                            $agrementProvenance = null === $agrementDate ? null : $decision->getProvenance();
                        }
                        break;
                    default:
                        break;
                }
                continue;
            }

            if (Deliberation::Retrait === $deliberation) {
                $retraitInitial = $decision->getDateCommission();
            }

            if (null !== $retraitInitial) {
                if (Deliberation::Accord === $deliberation
                    || ($conseilEtat && null === $retrait && Deliberation::Indetermine === $deliberation)) {
                    $retrait = null;
                } else {
                    $retrait = $retraitInitial;
                }
            }
        }

        return new AgrementResult($agrementDate, $agrementProvenance, $retrait);
    }

    public function actualiser(Demande $demande): AgrementResult
    {
        $decisions = $demande->getDecisions()->toArray();
        usort($decisions, static fn (Decision $a, Decision $b) => [$a->getDateCommission()?->getTimestamp() ?? PHP_INT_MIN, $a->getId() ?? PHP_INT_MAX]
            <=> [$b->getDateCommission()?->getTimestamp() ?? PHP_INT_MIN, $b->getId() ?? PHP_INT_MAX]);

        $resultat = $this->calculer($decisions);
        $demande->setDateAgrement($resultat->dateAgrement)
            ->setProvenanceAgrement($resultat->provenance)
            ->setDateRetrait($resultat->dateRetrait);

        return $resultat;
    }
}

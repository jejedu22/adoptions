<?php

namespace App\Service;

use App\Entity\Decision;
use App\Entity\Demande;
use App\Entity\Enum\CadreIntervention;
use App\Entity\Enum\Deliberation;

/** Déduit le statut, la frise d'avancement et la prochaine étape d'une demande à partir de ses données. */
final class StatutCalculator
{
    /** Délai (en mois) après lequel un agrément doit être réactualisé (règle de l'ancienne application). */
    public const MOIS_REACTUALISATION = 20;

    public function calculer(Demande $d, ?\DateTimeImmutable $aujourdhui = null): Statut
    {
        $aujourdhui ??= new \DateTimeImmutable('today');

        $rapports = array_values(array_filter(
            $d->getHistoriques()->toArray(),
            static fn ($h) => in_array($h->getCadre(), [CadreIntervention::DemandeInitiale->value, CadreIntervention::Reprise->value, CadreIntervention::SecondeEvaluation->value], true),
        ));
        $attendus = max(2, count($rapports));
        $recus = count(array_filter($rapports, static fn ($h) => null !== $h->getDateRecu()));
        $commissions = array_values(array_filter($d->getDecisions()->toArray(), static fn (Decision $x) => Decision::isInitiale($x)));
        $commissionPrise = [] !== array_filter($commissions, static fn (Decision $x) => Deliberation::Indetermine->value !== $x->getDeliberation());
        $prochaineCommission = null;
        foreach ($commissions as $c) {
            $prochaineCommission ??= $c->getDateDecisionPrevue() ?? $c->getDateCommission();
        }
        $recoursEnCours = [] !== array_filter($d->getDecisions()->toArray(), static fn (Decision $x) => null !== $x->getDateDepot() && Deliberation::Indetermine->value === $x->getDeliberation());
        $refus = [] !== array_filter($d->getDecisions()->toArray(), static fn (Decision $x) => in_array($x->getDeliberation(), [Deliberation::Refus->value, Deliberation::Retrait->value], true));

        $fmt = static fn (?\DateTimeImmutable $x): string => $x ? $x->format('d/m/Y') : '';
        $etapes = [
            ['libelle' => 'Demande', 'detail' => $fmt($d->getDateDemande()), 'fait' => null !== $d->getDateDemande()],
            ['libelle' => 'Dossier reçu', 'detail' => $fmt($d->getDateReceptionDossier()), 'fait' => null !== $d->getDateReceptionDossier()],
            ['libelle' => 'Rapports', 'detail' => sprintf('%d sur %d', min($recus, $attendus), $attendus), 'fait' => $recus >= $attendus],
            ['libelle' => 'Commission', 'detail' => $commissionPrise ? '' : ($prochaineCommission ? 'prévue '.$prochaineCommission->format('d/m') : ''), 'fait' => $commissionPrise],
            ['libelle' => 'Agrément', 'detail' => $fmt($d->getDateAgrement()), 'fait' => null !== $d->getDateAgrement()],
        ];
        $courante = false;
        foreach ($etapes as &$e) {
            $e['etat'] = $e['fait'] ? 'done' : (!$courante ? 'cur' : 'todo');
            $courante = $courante || !$e['fait'];
            unset($e['fait']);
        }
        unset($e);

        $aReactualiser = null !== $d->getDateAgrement()
            && $d->isAActualiser()
            && $d->getDateAgrement()->modify('+'.self::MOIS_REACTUALISATION.' months') < $aujourdhui;

        [$libelle, $ton, $prochaine] = match (true) {
            null !== $d->getDateSortie() => ['Sorti', Statut::MUTE, ''],
            null !== $d->getDateSuspension() => ['Suspendu', Statut::WARN, 'Reprendre ou clore l\'instruction'],
            null !== $d->getDateRetrait() => ['Retrait', Statut::BAD, $recoursEnCours ? 'Décision sur recours attendue' : ''],
            null !== $d->getDateAgrement() => $aReactualiser
                ? ['À réactualiser', Statut::INFO, 'Réactualiser la demande']
                : ['Agréé', Statut::OK, ''],
            $refus && $recoursEnCours => ['Refus · recours', Statut::BAD, 'Décision sur recours attendue'],
            $refus => ['Refus', Statut::BAD, ''],
            default => ['En instruction', Statut::WARN, self::prochaine($etapes, $d, $rapports)],
        };

        return new Statut($libelle, $ton, $etapes, $prochaine);
    }

    /** @param list<array{libelle: string, detail: string, etat: string}> $etapes */
    private static function prochaine(array $etapes, Demande $d, array $rapports): string
    {
        foreach ($etapes as $e) {
            if ('cur' !== $e['etat']) {
                continue;
            }

            return match ($e['libelle']) {
                'Dossier reçu' => 'Enregistrer la réception du dossier',
                'Rapports' => self::rapportAttendu($rapports),
                'Commission' => 'Saisir la décision de la commission',
                default => 'Instruire la demande',
            };
        }

        return '';
    }

    private static function rapportAttendu(array $rapports): string
    {
        foreach ($rapports as $h) {
            if (null === $h->getDateRecu()) {
                return sprintf('Rapport attendu : %s', $h->getIntervenant());
            }
        }

        return 'Demander les rapports d\'évaluation';
    }

    /** La demande la plus significative d'un dossier : la plus récente. */
    public static function demandeCourante(iterable $demandes): ?Demande
    {
        $courante = null;
        foreach ($demandes as $d) {
            if (null === $courante || ($d->getDateDemande()?->getTimestamp() ?? 0) >= ($courante->getDateDemande()?->getTimestamp() ?? 0)) {
                $courante = $d;
            }
        }

        return $courante;
    }
}

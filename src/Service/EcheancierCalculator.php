<?php

namespace App\Service;

use App\Entity\Demande;
use App\Repository\ParametreRepository;

/** Calcule les dates prévues de l'échéancier d'une demande à partir des délais paramétrés (en mois). */
final class EcheancierCalculator
{
    public const DEMANDE__RECEPTION_DOSSIER = 1;
    public const RECEPTION_DOSSIER__REPONSE_TRAVAILLEUR_SOCIAL = 2;
    public const RECEPTION_DOSSIER__REPONSE_PSYCHOLOGUE = 3;
    public const RECEPTION_DOSSIER__DECISION = 4;
    public const DEMANDE_RENOUVELLEMENT__REPONSE_TRAVAILLEUR_SOCIAL = 5;
    public const DEUX_SURVEILLANCES_AVANT_JUGEMENT = 6;
    public const RECOURS_ADMINISTRATIF__REPONSE_TRAVAILLEUR_SOCIAL = 7;
    public const RECOURS_ADMINISTRATIF__REPONSE_PSYCHOLOGUE = 8;

    public function __construct(private readonly ParametreRepository $parametres)
    {
    }

    /** @return array<string, ?\DateTimeImmutable> */
    public function calculer(Demande $demande): array
    {
        $mois = $this->parametres->moisParId();
        $ajouter = static fn (?\DateTimeImmutable $d, int $n): ?\DateTimeImmutable => $d?->modify(sprintf('+%d months', $n));

        $reception = $demande->getDateReceptionDossier() ?? $ajouter($demande->getDateDemande(), $mois[self::DEMANDE__RECEPTION_DOSSIER] ?? 0);

        return [
            'Réception du dossier' => $reception,
            'Réponse du travailleur social' => $ajouter($reception, $mois[self::RECEPTION_DOSSIER__REPONSE_TRAVAILLEUR_SOCIAL] ?? 0),
            'Réponse du psychologue' => $ajouter($reception, $mois[self::RECEPTION_DOSSIER__REPONSE_PSYCHOLOGUE] ?? 0),
            'Décision de la commission' => $ajouter($reception, $mois[self::RECEPTION_DOSSIER__DECISION] ?? 0),
        ];
    }
}

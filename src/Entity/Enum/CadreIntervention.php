<?php

namespace App\Entity\Enum;

enum CadreIntervention: int
{
    case Indetermine = 0;
    case DemandeInitiale = 1;
    case SecondeEvaluation = 2;
    case Reprise = 3;
    case RecoursAdministratif = 4;
    case SurveillancePlacement = 5;

    public function label(): string
    {
        return match ($this) {
            self::Indetermine => 'Indéterminé',
            self::DemandeInitiale => 'Demande initiale',
            self::SecondeEvaluation => 'Seconde évaluation',
            self::Reprise => 'Reprise',
            self::RecoursAdministratif => 'Recours administratif',
            self::SurveillancePlacement => 'Surveillance de placement',
        };
    }
}

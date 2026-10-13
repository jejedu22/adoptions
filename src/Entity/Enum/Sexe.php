<?php

namespace App\Entity\Enum;

enum Sexe: int
{
    case Indetermine = 1;
    case Masculin = 2;
    case Feminin = 3;

    public function label(): string
    {
        return match ($this) {
            self::Indetermine => 'Indéterminé',
            self::Masculin => 'Masculin',
            self::Feminin => 'Féminin',
        };
    }
}

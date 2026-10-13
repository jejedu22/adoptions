<?php

namespace App\Entity\Enum;

enum Deliberation: int
{
    case Indetermine = 0;
    case Refus = 1;
    case Retrait = 2;
    case Accord = 3;

    public function label(): string
    {
        return match ($this) {
            self::Indetermine => 'Indéterminée',
            self::Refus => 'Refus',
            self::Retrait => 'Retrait',
            self::Accord => 'Accord',
        };
    }
}

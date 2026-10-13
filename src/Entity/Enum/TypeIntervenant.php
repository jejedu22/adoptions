<?php

namespace App\Entity\Enum;

enum TypeIntervenant: string
{
    case Psychologue = 'psychologue';
    case AssistanteSociale = 'assistante_sociale';

    public function label(): string
    {
        return match ($this) {
            self::Psychologue => 'Psychologue',
            self::AssistanteSociale => 'Travailleur social',
        };
    }
}

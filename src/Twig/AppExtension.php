<?php

namespace App\Twig;

use App\Entity\Enum\CadreIntervention;
use App\Entity\Enum\Deliberation;
use App\Entity\Enum\Sexe;
use App\Service\Catalogue;
use Twig\Extension\AbstractExtension;
use Twig\TwigFunction;

class AppExtension extends AbstractExtension
{
    public function getFunctions(): array
    {
        return [
            new TwigFunction('menus', Catalogue::menus(...)),
            new TwigFunction('enum_label', $this->enumLabel(...)),
        ];
    }

    public function enumLabel(int $value, string $enum = 'deliberation'): string
    {
        $class = match ($enum) {
            'cadre' => CadreIntervention::class,
            'sexe' => Sexe::class,
            default => Deliberation::class,
        };

        return $class::tryFrom($value)?->label() ?? '';
    }
}

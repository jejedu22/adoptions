<?php

namespace App\Twig;

use App\Entity\Enum\CadreIntervention;
use App\Entity\Enum\Deliberation;
use App\Entity\Enum\Sexe;
use App\Service\Catalogue;
use App\Service\TacheFinder;
use Twig\Extension\AbstractExtension;
use Twig\TwigFunction;

class AppExtension extends AbstractExtension
{
    public function __construct(private readonly TacheFinder $taches)
    {
    }

    public function getFunctions(): array
    {
        return [
            new TwigFunction('menus', Catalogue::menus(...)),
            new TwigFunction('enum_label', $this->enumLabel(...)),
            new TwigFunction('nb_retards', $this->taches->nombreEnRetard(...)),
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

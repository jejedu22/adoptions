<?php

namespace App\Import;

/** Résultat (ou plan, en mode --check) d'une étape d'import. */
final class StepReport
{
    public bool $tableAbsente = false;
    public string $table = '';
    /** @var list<string> */
    public array $colonnes = [];
    /** @var list<string> */
    public array $colonnesAbsentes = [];
    public int $nbLignes = 0;
    public int $lues = 0;
    public int $crees = 0;
    public int $misAJour = 0;
    public int $ignorees = 0;
    public int $refsIntrouvables = 0;

    public function __construct(public readonly Step $step)
    {
    }
}

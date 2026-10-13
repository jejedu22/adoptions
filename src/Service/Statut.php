<?php

namespace App\Service;

/** Statut synthétique d'une demande, tel qu'affiché dans les listes et sur la fiche dossier. */
final class Statut
{
    public const OK = 'ok';
    public const WARN = 'warn';
    public const BAD = 'bad';
    public const INFO = 'info';
    public const MUTE = 'mute';

    /**
     * @param list<array{libelle: string, detail: string, etat: 'done'|'cur'|'todo'}> $etapes
     */
    public function __construct(
        public readonly string $libelle,
        public readonly string $ton,
        public readonly array $etapes,
        public readonly string $prochaineEtape,
    ) {
    }
}

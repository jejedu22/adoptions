<?php

namespace App\Import;

/**
 * Une étape d'import : une table de l'ancienne base vers une entité.
 *
 *  - key      : colonne(s) identifiant la ligne (sinon une empreinte de la ligne est utilisée)
 *  - fields   : propriété => 'COLONNE' ou ['COLONNE', 'transformation']
 *  - refs     : propriété => ['COLONNE', Entité::class]   (clé étrangère vers une entité déjà importée)
 *  - defaults : propriété => valeur fixe
 *  - requires : colonnes qui doivent être renseignées pour que la ligne soit importée
 *  - map      : fonction (ligne, contexte) => propriétés supplémentaires (ou null pour ignorer la ligne)
 *  - after    : fonction (entité, ligne, contexte) appelée après hydratation
 */
final class Step
{
    /**
     * @param class-string                                  $entity
     * @param list<string>                                  $key
     * @param array<string, string|array{0: string, 1: string}> $fields
     * @param array<string, array{0: string, 1: class-string}>  $refs
     * @param array<string, mixed>                          $defaults
     * @param list<string>                                  $requires
     */
    public function __construct(
        public readonly string $name,
        public readonly string $table,
        public readonly string $entity,
        public readonly array $key = [],
        public readonly array $fields = [],
        public readonly array $refs = [],
        public readonly array $defaults = [],
        public readonly array $requires = [],
        public readonly ?\Closure $map = null,
        public readonly ?\Closure $after = null,
        public readonly string $mapKey = '',
    ) {
    }

    /** Clé de correspondance : plusieurs étapes peuvent alimenter la même entité à partir de tables différentes. */
    public function scope(): string
    {
        return $this->mapKey ?: $this->entity;
    }
}

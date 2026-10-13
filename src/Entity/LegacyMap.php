<?php

namespace App\Entity;

use Doctrine\ORM\Mapping as ORM;

/** Correspondance entre un enregistrement de l'ancienne base et l'entité créée, pour rendre l'import rejouable. */
#[ORM\Entity]
#[ORM\Table(name: 'legacy_map')]
#[ORM\UniqueConstraint(columns: ['entity', 'legacy_key'])]
class LegacyMap
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    public function __construct(
        #[ORM\Column(length: 150)]
        private string $entity,
        #[ORM\Column(length: 191)]
        private string $legacyKey,
        #[ORM\Column]
        private int $newId,
    ) {
    }

    public function getEntity(): string { return $this->entity; }

    public function getLegacyKey(): string { return $this->legacyKey; }

    public function getNewId(): int { return $this->newId; }
}

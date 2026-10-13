<?php

namespace App\Import;

use App\Entity\Intervenant;
use Doctrine\DBAL\Connection;
use Doctrine\ORM\EntityManagerInterface;

/** État partagé par les étapes d'un import : correspondances ancien/nouvel identifiant, tables de référence en mémoire. */
final class ImportContext
{
    /** @var array<string, int> "portée|clé" => identifiant de la nouvelle entité */
    public array $map = [];

    /** @var array<string, array<string, array<string, mixed>>> table => colonne => valeur => première ligne */
    private array $index = [];

    /** @var array<int, array{psy: bool, as: bool}> */
    private array $typesIntervenants = [];

    /** @param array<string, string> $tables nom en minuscules => nom réel dans la base source */
    public function __construct(
        public readonly EntityManagerInterface $em,
        public readonly Connection $source,
        public readonly array $tables,
    ) {
    }

    public function resolve(string $class, string $key): ?object
    {
        $id = $this->map[$class.'|'.$key] ?? null;

        return null === $id ? null : $this->em->getReference($class, $id);
    }

    /** Retourne la première ligne de $table dont $colonne vaut $valeur (table chargée une fois en mémoire). */
    public function ligne(string $table, string $colonne, mixed $valeur): ?array
    {
        $table = strtolower($table);
        if (null === $valeur || !isset($this->tables[$table])) {
            return null;
        }
        if (!isset($this->index[$table][$colonne])) {
            $this->index[$table][$colonne] = [];
            foreach ($this->source->iterateAssociative('SELECT * FROM '.$this->source->quoteSingleIdentifier($this->tables[$table])) as $row) {
                $row = array_change_key_case($row, CASE_LOWER);
                if (isset($row[$colonne])) {
                    $this->index[$table][$colonne][(string) $row[$colonne]] ??= $row;
                }
            }
        }

        return $this->index[$table][$colonne][(string) $valeur] ?? null;
    }

    public function marquerType(?Intervenant $intervenant, bool $psychologue): void
    {
        if (null !== $intervenant && null !== $intervenant->getId()) {
            $this->typesIntervenants[$intervenant->getId()][$psychologue ? 'psy' : 'as'] = true;
        }
    }

    /** @return array<int, array{psy?: bool, as?: bool}> */
    public function typesIntervenants(): array
    {
        return $this->typesIntervenants;
    }
}

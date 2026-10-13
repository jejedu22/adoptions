<?php

namespace App\Service;

use App\Entity\Demande;

/** Une action à mener sur une demande (rapport en retard, décision à saisir, réactualisation, surveillance…). */
final class Tache
{
    public function __construct(
        public readonly string $type,
        public readonly string $libelle,
        public readonly string $detail,
        public readonly Demande $demande,
        public readonly ?\DateTimeImmutable $echeance,
        /** Jours de retard (positif) ou restants avant échéance (négatif) */
        public readonly int $retard,
        public readonly ?int $interventionId = null,
        public readonly ?string $intervenant = null,
    ) {
    }

    public function enRetard(): bool
    {
        return $this->retard > 0;
    }

    public function ton(): string
    {
        return $this->enRetard() ? Statut::BAD : ($this->retard >= -7 ? Statut::WARN : Statut::INFO);
    }

    public function etiquette(): string
    {
        return match (true) {
            $this->retard > 0 => 'En retard · '.$this->retard.' j',
            0 === $this->retard => 'Aujourd\'hui',
            $this->retard >= -7 => 'Dans '.-$this->retard.' j',
            default => 'Dans '.-$this->retard.' j',
        };
    }
}

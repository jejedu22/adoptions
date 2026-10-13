<?php

namespace App\Entity;

use Doctrine\Common\Collections\ArrayCollection;
use Doctrine\Common\Collections\Collection;
use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity]
class Intervenant implements \Stringable
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    #[ORM\Column(length: 100, nullable: false)]
    private string $nom = '';

    #[ORM\Column(length: 20, nullable: true)]
    private ?string $titre = null;

    #[ORM\Column(length: 30, nullable: true)]
    private ?string $telephone = null;

    #[ORM\Column]
    private bool $actif = false;

    #[ORM\Column(length: 30)]
    private string $type = 'assistante_sociale';

    #[ORM\ManyToOne(targetEntity: Qualite::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?Qualite $qualite = null;

    #[ORM\ManyToOne(targetEntity: Institution::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?Institution $institution = null;

    #[ORM\ManyToOne(targetEntity: Permanence::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?Permanence $permanence = null;

    public function getId(): ?int { return $this->id; }

    public function getNom(): string { return $this->nom; }

    public function setNom(string $nom): static { $this->nom = $nom; return $this; }

    public function getTitre(): ?string { return $this->titre; }

    public function setTitre(?string $titre): static { $this->titre = $titre; return $this; }

    public function getTelephone(): ?string { return $this->telephone; }

    public function setTelephone(?string $telephone): static { $this->telephone = $telephone; return $this; }

    public function isActif(): bool { return $this->actif; }

    public function setActif(bool $actif): static { $this->actif = $actif; return $this; }

    public function getType(): string { return $this->type; }

    public function setType(string $type): static { $this->type = $type; return $this; }

    public function getQualite(): ?Qualite { return $this->qualite; }

    public function setQualite(?Qualite $qualite): static { $this->qualite = $qualite; return $this; }

    public function getInstitution(): ?Institution { return $this->institution; }

    public function setInstitution(?Institution $institution): static { $this->institution = $institution; return $this; }

    public function getPermanence(): ?Permanence { return $this->permanence; }

    public function setPermanence(?Permanence $permanence): static { $this->permanence = $permanence; return $this; }

    public function __toString(): string { return (string) $this->nom; }
}

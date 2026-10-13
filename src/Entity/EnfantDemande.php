<?php

namespace App\Entity;

use Doctrine\Common\Collections\ArrayCollection;
use Doctrine\Common\Collections\Collection;
use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity]
class EnfantDemande
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    #[ORM\ManyToOne(targetEntity: Demande::class, inversedBy: 'enfants')]
    #[ORM\JoinColumn(nullable: false, onDelete: 'CASCADE')]
    private Demande $demande;

    #[ORM\Column]
    private int $sexe = 1;

    #[ORM\Column(nullable: true)]
    private ?int $ageMax = null;

    #[ORM\Column]
    private bool $handicap = false;

    #[ORM\ManyToOne(targetEntity: OrigineEthnique::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?OrigineEthnique $origineEthnique = null;

    public function getId(): ?int { return $this->id; }

    public function getDemande(): Demande { return $this->demande; }

    public function setDemande(Demande $demande): static { $this->demande = $demande; return $this; }

    public function getSexe(): int { return $this->sexe; }

    public function setSexe(int $sexe): static { $this->sexe = $sexe; return $this; }

    public function getAgeMax(): ?int { return $this->ageMax; }

    public function setAgeMax(?int $ageMax): static { $this->ageMax = $ageMax; return $this; }

    public function isHandicap(): bool { return $this->handicap; }

    public function setHandicap(bool $handicap): static { $this->handicap = $handicap; return $this; }

    public function getOrigineEthnique(): ?OrigineEthnique { return $this->origineEthnique; }

    public function setOrigineEthnique(?OrigineEthnique $origineEthnique): static { $this->origineEthnique = $origineEthnique; return $this; }
}

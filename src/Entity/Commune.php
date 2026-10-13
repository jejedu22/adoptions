<?php

namespace App\Entity;

use Doctrine\Common\Collections\ArrayCollection;
use Doctrine\Common\Collections\Collection;
use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity]
class Commune implements \Stringable
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    #[ORM\Column(length: 150, nullable: false)]
    private string $nom = '';

    #[ORM\Column(length: 10, nullable: true)]
    private ?string $codePostal = null;

    #[ORM\Column(length: 10, nullable: true)]
    private ?string $codeInsee = null;

    #[ORM\ManyToOne(targetEntity: Canton::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?Canton $canton = null;

    #[ORM\ManyToOne(targetEntity: Circonscription::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?Circonscription $circonscription = null;

    public function getId(): ?int { return $this->id; }

    public function getNom(): string { return $this->nom; }

    public function setNom(string $nom): static { $this->nom = $nom; return $this; }

    public function getCodePostal(): ?string { return $this->codePostal; }

    public function setCodePostal(?string $codePostal): static { $this->codePostal = $codePostal; return $this; }

    public function getCodeInsee(): ?string { return $this->codeInsee; }

    public function setCodeInsee(?string $codeInsee): static { $this->codeInsee = $codeInsee; return $this; }

    public function getCanton(): ?Canton { return $this->canton; }

    public function setCanton(?Canton $canton): static { $this->canton = $canton; return $this; }

    public function getCirconscription(): ?Circonscription { return $this->circonscription; }

    public function setCirconscription(?Circonscription $circonscription): static { $this->circonscription = $circonscription; return $this; }

    public function __toString(): string { return (string) $this->nom; }
}

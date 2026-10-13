<?php

namespace App\Entity;

use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity]
class Circonscription implements \Stringable
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    #[ORM\Column(length: 150)]
    #[Assert\NotBlank]
    private string $nom = '';

    #[ORM\Column(length: 200, nullable: true)]
    private ?string $adresse = null;

    #[ORM\Column(length: 10, nullable: true)]
    private ?string $codePostal = null;

    #[ORM\Column(length: 100, nullable: true)]
    private ?string $ville = null;

    #[ORM\Column(length: 150, nullable: true)]
    private ?string $responsable = null;

    public function getId(): ?int { return $this->id; }

    public function getNom(): string { return $this->nom; }

    public function setNom(string $nom): static { $this->nom = $nom; return $this; }

    public function getAdresse(): ?string { return $this->adresse; }

    public function setAdresse(?string $adresse): static { $this->adresse = $adresse; return $this; }

    public function getCodePostal(): ?string { return $this->codePostal; }

    public function setCodePostal(?string $codePostal): static { $this->codePostal = $codePostal; return $this; }

    public function getVille(): ?string { return $this->ville; }

    public function setVille(?string $ville): static { $this->ville = $ville; return $this; }

    public function getResponsable(): ?string { return $this->responsable; }

    public function setResponsable(?string $responsable): static { $this->responsable = $responsable; return $this; }

    public function __toString(): string { return (string) $this->nom; }
}

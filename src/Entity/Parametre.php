<?php

namespace App\Entity;

use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity(repositoryClass: \App\Repository\ParametreRepository::class)]
class Parametre implements \Stringable
{
    #[ORM\Id, ORM\Column]
    private ?int $id = null;

    #[ORM\Column(length: 150)]
    #[Assert\NotBlank]
    private string $libelle = '';

    #[ORM\Column]
    #[Assert\PositiveOrZero]
    private int $nombreMois = 0;

    public function getId(): ?int { return $this->id; }

    public function setId(int $id): static { $this->id = $id; return $this; }

    public function getLibelle(): string { return $this->libelle; }

    public function setLibelle(string $libelle): static { $this->libelle = $libelle; return $this; }

    public function getNombreMois(): int { return $this->nombreMois; }

    public function setNombreMois(int $nombreMois): static { $this->nombreMois = $nombreMois; return $this; }

    public function __toString(): string { return (string) $this->libelle; }
}

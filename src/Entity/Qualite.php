<?php

namespace App\Entity;

use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity]
#[ORM\Table(name: 'qualite')]
class Qualite implements \Stringable
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    #[ORM\Column(length: 150)]
    #[Assert\NotBlank]
    private string $libelle = '';

    public function getId(): ?int { return $this->id; }

    public function getLibelle(): string { return $this->libelle; }

    public function setLibelle(string $libelle): static { $this->libelle = $libelle; return $this; }

    public function __toString(): string { return (string) $this->libelle; }
}

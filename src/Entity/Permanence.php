<?php

namespace App\Entity;

use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity]
class Permanence implements \Stringable
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    #[ORM\Column(length: 150)]
    #[Assert\NotBlank]
    private string $lieu = '';

    #[ORM\Column(length: 100, nullable: true)]
    private ?string $ville = null;

    public function getId(): ?int { return $this->id; }

    public function getLieu(): string { return $this->lieu; }

    public function setLieu(string $lieu): static { $this->lieu = $lieu; return $this; }

    public function getVille(): ?string { return $this->ville; }

    public function setVille(?string $ville): static { $this->ville = $ville; return $this; }

    public function __toString(): string { return (string) $this->lieu; }
}

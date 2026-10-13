<?php

namespace App\Entity;

use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity]
class Tribunal implements \Stringable
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    #[ORM\Column(length: 150)]
    #[Assert\NotBlank]
    private string $libelle = '';

    #[ORM\Column(length: 150, nullable: true)]
    private ?string $adresse1 = null;

    #[ORM\Column(length: 150, nullable: true)]
    private ?string $adresse2 = null;

    #[ORM\Column(length: 10, nullable: true)]
    private ?string $codePostal = null;

    #[ORM\Column(length: 100, nullable: true)]
    private ?string $ville = null;

    #[ORM\Column(type: 'text', nullable: true)]
    private ?string $commentaires = null;

    public function getId(): ?int { return $this->id; }

    public function getLibelle(): string { return $this->libelle; }

    public function setLibelle(string $libelle): static { $this->libelle = $libelle; return $this; }

    public function getAdresse1(): ?string { return $this->adresse1; }

    public function setAdresse1(?string $adresse1): static { $this->adresse1 = $adresse1; return $this; }

    public function getAdresse2(): ?string { return $this->adresse2; }

    public function setAdresse2(?string $adresse2): static { $this->adresse2 = $adresse2; return $this; }

    public function getCodePostal(): ?string { return $this->codePostal; }

    public function setCodePostal(?string $codePostal): static { $this->codePostal = $codePostal; return $this; }

    public function getVille(): ?string { return $this->ville; }

    public function setVille(?string $ville): static { $this->ville = $ville; return $this; }

    public function getCommentaires(): ?string { return $this->commentaires; }

    public function setCommentaires(?string $commentaires): static { $this->commentaires = $commentaires; return $this; }

    public function __toString(): string { return (string) $this->libelle; }
}

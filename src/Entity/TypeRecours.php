<?php

namespace App\Entity;

use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity]
#[ORM\Table(name: 'type_recours')]
class TypeRecours implements \Stringable
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    #[ORM\Column(length: 150)]
    #[Assert\NotBlank]
    private string $libelle = '';

    #[ORM\Column]
    private bool $conseilEtat = false;

    public function getId(): ?int { return $this->id; }

    public function getLibelle(): string { return $this->libelle; }

    public function setLibelle(string $libelle): static { $this->libelle = $libelle; return $this; }

    public function isConseilEtat(): bool { return $this->conseilEtat; }

    public function setConseilEtat(bool $conseilEtat): static { $this->conseilEtat = $conseilEtat; return $this; }

    public function __toString(): string { return (string) $this->libelle; }
}

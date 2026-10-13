<?php

namespace App\Entity;

use Doctrine\Common\Collections\ArrayCollection;
use Doctrine\Common\Collections\Collection;
use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity]
class Surveillance
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    #[ORM\ManyToOne(targetEntity: Placement::class, inversedBy: 'surveillances')]
    #[ORM\JoinColumn(nullable: false, onDelete: 'CASCADE')]
    private Placement $placement;

    #[ORM\ManyToOne(targetEntity: Intervenant::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?Intervenant $intervenant = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $datePrevue = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateRealisee = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateEnvoiTgi = null;

    #[ORM\Column(type: 'text', nullable: true)]
    private ?string $remarques = null;

    public function getId(): ?int { return $this->id; }

    public function getPlacement(): Placement { return $this->placement; }

    public function setPlacement(Placement $placement): static { $this->placement = $placement; return $this; }

    public function getIntervenant(): ?Intervenant { return $this->intervenant; }

    public function setIntervenant(?Intervenant $intervenant): static { $this->intervenant = $intervenant; return $this; }

    public function getDatePrevue(): ?\DateTimeImmutable { return $this->datePrevue; }

    public function setDatePrevue(?\DateTimeImmutable $datePrevue): static { $this->datePrevue = $datePrevue; return $this; }

    public function getDateRealisee(): ?\DateTimeImmutable { return $this->dateRealisee; }

    public function setDateRealisee(?\DateTimeImmutable $dateRealisee): static { $this->dateRealisee = $dateRealisee; return $this; }

    public function getDateEnvoiTgi(): ?\DateTimeImmutable { return $this->dateEnvoiTgi; }

    public function setDateEnvoiTgi(?\DateTimeImmutable $dateEnvoiTgi): static { $this->dateEnvoiTgi = $dateEnvoiTgi; return $this; }

    public function getRemarques(): ?string { return $this->remarques; }

    public function setRemarques(?string $remarques): static { $this->remarques = $remarques; return $this; }
}

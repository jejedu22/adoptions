<?php

namespace App\Entity;

use Doctrine\Common\Collections\ArrayCollection;
use Doctrine\Common\Collections\Collection;
use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity]
class HistoriqueIntervenant
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    #[ORM\ManyToOne(targetEntity: Demande::class, inversedBy: 'historiques')]
    #[ORM\JoinColumn(nullable: false, onDelete: 'CASCADE')]
    private Demande $demande;

    #[ORM\ManyToOne(targetEntity: Intervenant::class)]
    #[ORM\JoinColumn(nullable: false)]
    private Intervenant $intervenant;

    #[ORM\Column]
    private int $cadre = 1;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateDemande = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $datePrevue = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateRecu = null;

    public function getId(): ?int { return $this->id; }

    public function getDemande(): Demande { return $this->demande; }

    public function setDemande(Demande $demande): static { $this->demande = $demande; return $this; }

    public function getIntervenant(): Intervenant { return $this->intervenant; }

    public function setIntervenant(Intervenant $intervenant): static { $this->intervenant = $intervenant; return $this; }

    public function getCadre(): int { return $this->cadre; }

    public function setCadre(int $cadre): static { $this->cadre = $cadre; return $this; }

    public function getDateDemande(): ?\DateTimeImmutable { return $this->dateDemande; }

    public function setDateDemande(?\DateTimeImmutable $dateDemande): static { $this->dateDemande = $dateDemande; return $this; }

    public function getDatePrevue(): ?\DateTimeImmutable { return $this->datePrevue; }

    public function setDatePrevue(?\DateTimeImmutable $datePrevue): static { $this->datePrevue = $datePrevue; return $this; }

    public function getDateRecu(): ?\DateTimeImmutable { return $this->dateRecu; }

    public function setDateRecu(?\DateTimeImmutable $dateRecu): static { $this->dateRecu = $dateRecu; return $this; }
}

<?php

namespace App\Entity;

use Doctrine\Common\Collections\ArrayCollection;
use Doctrine\Common\Collections\Collection;
use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity]
class Decision
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    #[ORM\ManyToOne(targetEntity: Demande::class, inversedBy: 'decisions')]
    #[ORM\JoinColumn(nullable: false, onDelete: 'CASCADE')]
    private Demande $demande;

    #[ORM\Column]
    private int $provenance = -2;

    #[ORM\ManyToOne(targetEntity: TypeRecours::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?TypeRecours $typeRecours = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $commissionInitiale = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateCommission = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateSignature = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateDepot = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateDecisionPrevue = null;

    #[ORM\Column]
    private int $deliberation = 0;

    #[ORM\Column(type: 'text', nullable: true)]
    private ?string $remarques = null;

    public function getId(): ?int { return $this->id; }

    public function getDemande(): Demande { return $this->demande; }

    public function setDemande(Demande $demande): static { $this->demande = $demande; return $this; }

    public function getProvenance(): int { return $this->provenance; }

    public function setProvenance(int $provenance): static { $this->provenance = $provenance; return $this; }

    public function getTypeRecours(): ?TypeRecours { return $this->typeRecours; }

    public function setTypeRecours(?TypeRecours $typeRecours): static { $this->typeRecours = $typeRecours; return $this; }

    public function getCommissionInitiale(): ?\DateTimeImmutable { return $this->commissionInitiale; }

    public function setCommissionInitiale(?\DateTimeImmutable $commissionInitiale): static { $this->commissionInitiale = $commissionInitiale; return $this; }

    public function getDateCommission(): ?\DateTimeImmutable { return $this->dateCommission; }

    public function setDateCommission(?\DateTimeImmutable $dateCommission): static { $this->dateCommission = $dateCommission; return $this; }

    public function getDateSignature(): ?\DateTimeImmutable { return $this->dateSignature; }

    public function setDateSignature(?\DateTimeImmutable $dateSignature): static { $this->dateSignature = $dateSignature; return $this; }

    public function getDateDepot(): ?\DateTimeImmutable { return $this->dateDepot; }

    public function setDateDepot(?\DateTimeImmutable $dateDepot): static { $this->dateDepot = $dateDepot; return $this; }

    public function getDateDecisionPrevue(): ?\DateTimeImmutable { return $this->dateDecisionPrevue; }

    public function setDateDecisionPrevue(?\DateTimeImmutable $dateDecisionPrevue): static { $this->dateDecisionPrevue = $dateDecisionPrevue; return $this; }

    public function getDeliberation(): int { return $this->deliberation; }

    public function setDeliberation(int $deliberation): static { $this->deliberation = $deliberation; return $this; }

    public function getRemarques(): ?string { return $this->remarques; }

    public function setRemarques(?string $remarques): static { $this->remarques = $remarques; return $this; }

    public static function isInitiale(self $decision): bool
    {
        return -2 === $decision->getProvenance();
    }
}

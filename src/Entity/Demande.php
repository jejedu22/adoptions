<?php

namespace App\Entity;

use Doctrine\Common\Collections\ArrayCollection;
use Doctrine\Common\Collections\Collection;
use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity]
class Demande
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    #[ORM\ManyToOne(targetEntity: Dossier::class, inversedBy: 'demandes')]
    #[ORM\JoinColumn(nullable: false, onDelete: 'CASCADE')]
    private Dossier $dossier;

    #[ORM\Column]
    private int $numero = 0;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateDemande = null;

    #[ORM\Column(nullable: true)]
    private ?int $nombreEnfants = null;

    #[ORM\ManyToOne(targetEntity: OrigineEthnique::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?OrigineEthnique $origineEthnique = null;

    #[ORM\Column]
    private bool $fratrie = false;

    #[ORM\Column]
    private bool $assistanteMaternelle = false;

    #[ORM\ManyToOne(targetEntity: NatureDemande::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?NatureDemande $nature = null;

    #[ORM\ManyToOne(targetEntity: TypeDemande::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?TypeDemande $type = null;

    #[ORM\Column(type: 'text', nullable: true)]
    private ?string $remarques = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateSortie = null;

    #[ORM\Column(length: 255, nullable: true)]
    private ?string $motifSortie = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateSuspension = null;

    #[ORM\Column(length: 255, nullable: true)]
    private ?string $motifSuspension = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateAgrement = null;

    #[ORM\Column(nullable: true)]
    private ?int $provenanceAgrement = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateRetrait = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $datePrevueDecision = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateReceptionDossier = null;

    #[ORM\Column]
    private bool $aActualiser = false;

    /** @var Collection<int, Decision> */
    #[ORM\OneToMany(targetEntity: Decision::class, mappedBy: 'demande', cascade: ['persist','remove'], orphanRemoval: true)]
    #[ORM\OrderBy(['commissionInitiale' => 'ASC', 'provenance' => 'ASC', 'id' => 'ASC'])]
    private Collection $decisions;

    /** @var Collection<int, HistoriqueIntervenant> */
    #[ORM\OneToMany(targetEntity: HistoriqueIntervenant::class, mappedBy: 'demande', cascade: ['persist','remove'], orphanRemoval: true)]
    #[ORM\OrderBy(['id' => 'ASC'])]
    private Collection $historiques;

    /** @var Collection<int, EnfantDemande> */
    #[ORM\OneToMany(targetEntity: EnfantDemande::class, mappedBy: 'demande', cascade: ['persist','remove'], orphanRemoval: true)]
    private Collection $enfants;

    /** @var Collection<int, Placement> */
    #[ORM\OneToMany(targetEntity: Placement::class, mappedBy: 'demande', cascade: ['persist','remove'], orphanRemoval: true)]
    private Collection $placements;

    public function __construct()
    {
        $this->decisions = new ArrayCollection();
        $this->historiques = new ArrayCollection();
        $this->enfants = new ArrayCollection();
        $this->placements = new ArrayCollection();
    }

    public function getId(): ?int { return $this->id; }

    public function getDossier(): Dossier { return $this->dossier; }

    public function setDossier(Dossier $dossier): static { $this->dossier = $dossier; return $this; }

    public function getNumero(): int { return $this->numero; }

    public function setNumero(int $numero): static { $this->numero = $numero; return $this; }

    public function getDateDemande(): ?\DateTimeImmutable { return $this->dateDemande; }

    public function setDateDemande(?\DateTimeImmutable $dateDemande): static { $this->dateDemande = $dateDemande; return $this; }

    public function getNombreEnfants(): ?int { return $this->nombreEnfants; }

    public function setNombreEnfants(?int $nombreEnfants): static { $this->nombreEnfants = $nombreEnfants; return $this; }

    public function getOrigineEthnique(): ?OrigineEthnique { return $this->origineEthnique; }

    public function setOrigineEthnique(?OrigineEthnique $origineEthnique): static { $this->origineEthnique = $origineEthnique; return $this; }

    public function isFratrie(): bool { return $this->fratrie; }

    public function setFratrie(bool $fratrie): static { $this->fratrie = $fratrie; return $this; }

    public function isAssistanteMaternelle(): bool { return $this->assistanteMaternelle; }

    public function setAssistanteMaternelle(bool $assistanteMaternelle): static { $this->assistanteMaternelle = $assistanteMaternelle; return $this; }

    public function getNature(): ?NatureDemande { return $this->nature; }

    public function setNature(?NatureDemande $nature): static { $this->nature = $nature; return $this; }

    public function getType(): ?TypeDemande { return $this->type; }

    public function setType(?TypeDemande $type): static { $this->type = $type; return $this; }

    public function getRemarques(): ?string { return $this->remarques; }

    public function setRemarques(?string $remarques): static { $this->remarques = $remarques; return $this; }

    public function getDateSortie(): ?\DateTimeImmutable { return $this->dateSortie; }

    public function setDateSortie(?\DateTimeImmutable $dateSortie): static { $this->dateSortie = $dateSortie; return $this; }

    public function getMotifSortie(): ?string { return $this->motifSortie; }

    public function setMotifSortie(?string $motifSortie): static { $this->motifSortie = $motifSortie; return $this; }

    public function getDateSuspension(): ?\DateTimeImmutable { return $this->dateSuspension; }

    public function setDateSuspension(?\DateTimeImmutable $dateSuspension): static { $this->dateSuspension = $dateSuspension; return $this; }

    public function getMotifSuspension(): ?string { return $this->motifSuspension; }

    public function setMotifSuspension(?string $motifSuspension): static { $this->motifSuspension = $motifSuspension; return $this; }

    public function getDateAgrement(): ?\DateTimeImmutable { return $this->dateAgrement; }

    public function setDateAgrement(?\DateTimeImmutable $dateAgrement): static { $this->dateAgrement = $dateAgrement; return $this; }

    public function getProvenanceAgrement(): ?int { return $this->provenanceAgrement; }

    public function setProvenanceAgrement(?int $provenanceAgrement): static { $this->provenanceAgrement = $provenanceAgrement; return $this; }

    public function getDateRetrait(): ?\DateTimeImmutable { return $this->dateRetrait; }

    public function setDateRetrait(?\DateTimeImmutable $dateRetrait): static { $this->dateRetrait = $dateRetrait; return $this; }

    public function getDatePrevueDecision(): ?\DateTimeImmutable { return $this->datePrevueDecision; }

    public function setDatePrevueDecision(?\DateTimeImmutable $datePrevueDecision): static { $this->datePrevueDecision = $datePrevueDecision; return $this; }

    public function getDateReceptionDossier(): ?\DateTimeImmutable { return $this->dateReceptionDossier; }

    public function setDateReceptionDossier(?\DateTimeImmutable $dateReceptionDossier): static { $this->dateReceptionDossier = $dateReceptionDossier; return $this; }

    public function isAActualiser(): bool { return $this->aActualiser; }

    public function setAActualiser(bool $aActualiser): static { $this->aActualiser = $aActualiser; return $this; }

    /** @return Collection<int, Decision> */
    public function getDecisions(): Collection { return $this->decisions; }

    public function addDecision(Decision $item): static
    {
        if (!$this->decisions->contains($item)) {
            $this->decisions->add($item);
            $item->setDemande($this);
        }

        return $this;
    }

    /** @return Collection<int, HistoriqueIntervenant> */
    public function getHistoriques(): Collection { return $this->historiques; }

    public function addHistorique(HistoriqueIntervenant $item): static
    {
        if (!$this->historiques->contains($item)) {
            $this->historiques->add($item);
            $item->setDemande($this);
        }

        return $this;
    }

    /** @return Collection<int, EnfantDemande> */
    public function getEnfants(): Collection { return $this->enfants; }

    public function addEnfant(EnfantDemande $item): static
    {
        if (!$this->enfants->contains($item)) {
            $this->enfants->add($item);
            $item->setDemande($this);
        }

        return $this;
    }

    /** @return Collection<int, Placement> */
    public function getPlacements(): Collection { return $this->placements; }

    public function addPlacement(Placement $item): static
    {
        if (!$this->placements->contains($item)) {
            $this->placements->add($item);
            $item->setDemande($this);
        }

        return $this;
    }
}

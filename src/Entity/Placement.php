<?php

namespace App\Entity;

use Doctrine\Common\Collections\ArrayCollection;
use Doctrine\Common\Collections\Collection;
use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity]
class Placement implements \Stringable
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    #[ORM\ManyToOne(targetEntity: Demande::class, inversedBy: 'placements')]
    #[ORM\JoinColumn(nullable: false, onDelete: 'CASCADE')]
    private Demande $demande;

    #[ORM\Column(length: 100, nullable: true)]
    private ?string $nomEnfant = null;

    #[ORM\Column(length: 100, nullable: true)]
    private ?string $prenomEnfant = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateNaissance = null;

    #[ORM\Column]
    private int $sexe = 1;

    #[ORM\ManyToOne(targetEntity: OrigineEthnique::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?OrigineEthnique $origineEthnique = null;

    #[ORM\ManyToOne(targetEntity: Oeuvre::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?Oeuvre $oeuvre = null;

    #[ORM\ManyToOne(targetEntity: Tribunal::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?Tribunal $tribunal = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $datePlacement = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateJugement = null;

    #[ORM\Column]
    private bool $pupille = false;

    #[ORM\Column(type: 'text', nullable: true)]
    private ?string $remarques = null;

    /** @var Collection<int, Surveillance> */
    #[ORM\OneToMany(targetEntity: Surveillance::class, mappedBy: 'placement', cascade: ['persist','remove'], orphanRemoval: true)]
    #[ORM\OrderBy(['datePrevue' => 'ASC'])]
    private Collection $surveillances;

    public function __construct()
    {
        $this->surveillances = new ArrayCollection();
    }

    public function getId(): ?int { return $this->id; }

    public function getDemande(): Demande { return $this->demande; }

    public function setDemande(Demande $demande): static { $this->demande = $demande; return $this; }

    public function getNomEnfant(): ?string { return $this->nomEnfant; }

    public function setNomEnfant(?string $nomEnfant): static { $this->nomEnfant = $nomEnfant; return $this; }

    public function getPrenomEnfant(): ?string { return $this->prenomEnfant; }

    public function setPrenomEnfant(?string $prenomEnfant): static { $this->prenomEnfant = $prenomEnfant; return $this; }

    public function getDateNaissance(): ?\DateTimeImmutable { return $this->dateNaissance; }

    public function setDateNaissance(?\DateTimeImmutable $dateNaissance): static { $this->dateNaissance = $dateNaissance; return $this; }

    public function getSexe(): int { return $this->sexe; }

    public function setSexe(int $sexe): static { $this->sexe = $sexe; return $this; }

    public function getOrigineEthnique(): ?OrigineEthnique { return $this->origineEthnique; }

    public function setOrigineEthnique(?OrigineEthnique $origineEthnique): static { $this->origineEthnique = $origineEthnique; return $this; }

    public function getOeuvre(): ?Oeuvre { return $this->oeuvre; }

    public function setOeuvre(?Oeuvre $oeuvre): static { $this->oeuvre = $oeuvre; return $this; }

    public function getTribunal(): ?Tribunal { return $this->tribunal; }

    public function setTribunal(?Tribunal $tribunal): static { $this->tribunal = $tribunal; return $this; }

    public function getDatePlacement(): ?\DateTimeImmutable { return $this->datePlacement; }

    public function setDatePlacement(?\DateTimeImmutable $datePlacement): static { $this->datePlacement = $datePlacement; return $this; }

    public function getDateJugement(): ?\DateTimeImmutable { return $this->dateJugement; }

    public function setDateJugement(?\DateTimeImmutable $dateJugement): static { $this->dateJugement = $dateJugement; return $this; }

    public function isPupille(): bool { return $this->pupille; }

    public function setPupille(bool $pupille): static { $this->pupille = $pupille; return $this; }

    public function getRemarques(): ?string { return $this->remarques; }

    public function setRemarques(?string $remarques): static { $this->remarques = $remarques; return $this; }

    /** @return Collection<int, Surveillance> */
    public function getSurveillances(): Collection { return $this->surveillances; }

    public function addSurveillance(Surveillance $item): static
    {
        if (!$this->surveillances->contains($item)) {
            $this->surveillances->add($item);
            $item->setPlacement($this);
        }

        return $this;
    }

    public function __toString(): string { return (string) $this->nomEnfant; }
}

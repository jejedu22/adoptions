<?php

namespace App\Entity;

use Doctrine\Common\Collections\ArrayCollection;
use Doctrine\Common\Collections\Collection;
use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity(repositoryClass: \App\Repository\DossierRepository::class)]
class Dossier implements \Stringable
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    #[ORM\Column(length: 20, nullable: false)]
    private string $numero = '';

    #[ORM\Column(length: 3, nullable: true)]
    private ?string $departement = null;

    #[ORM\Column(length: 30, nullable: true)]
    private ?string $telephone = null;

    #[ORM\Column(length: 150, nullable: true)]
    private ?string $adresse1 = null;

    #[ORM\Column(length: 150, nullable: true)]
    private ?string $adresse2 = null;

    #[ORM\Column(length: 10, nullable: true)]
    private ?string $codePostal = null;

    #[ORM\Column(length: 100, nullable: true)]
    private ?string $ville = null;

    #[ORM\Column]
    private bool $reinstruction = false;

    #[ORM\ManyToOne(targetEntity: Commune::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?Commune $commune = null;

    #[ORM\ManyToOne(targetEntity: Canton::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?Canton $canton = null;

    #[ORM\ManyToOne(targetEntity: Circonscription::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?Circonscription $circonscription = null;

    #[ORM\ManyToOne(targetEntity: SituationFamiliale::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?SituationFamiliale $situationFamiliale = null;

    #[ORM\Column(length: 150, nullable: true)]
    private ?string $regimeMatrimonial = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateMariage = null;

    #[ORM\Column(length: 100, nullable: true)]
    private ?string $lieuMariage = null;

    /** @var Collection<int, Individu> */
    #[ORM\OneToMany(targetEntity: Individu::class, mappedBy: 'dossier', cascade: ['persist','remove'], orphanRemoval: true)]
    private Collection $individus;

    /** @var Collection<int, Demande> */
    #[ORM\OneToMany(targetEntity: Demande::class, mappedBy: 'dossier', cascade: ['persist','remove'], orphanRemoval: true)]
    #[ORM\OrderBy(['dateDemande' => 'ASC'])]
    private Collection $demandes;

    public function __construct()
    {
        $this->individus = new ArrayCollection();
        $this->demandes = new ArrayCollection();
    }

    public function getId(): ?int { return $this->id; }

    public function getNumero(): string { return $this->numero; }

    public function setNumero(string $numero): static { $this->numero = $numero; return $this; }

    public function getDepartement(): ?string { return $this->departement; }

    public function setDepartement(?string $departement): static { $this->departement = $departement; return $this; }

    public function getTelephone(): ?string { return $this->telephone; }

    public function setTelephone(?string $telephone): static { $this->telephone = $telephone; return $this; }

    public function getAdresse1(): ?string { return $this->adresse1; }

    public function setAdresse1(?string $adresse1): static { $this->adresse1 = $adresse1; return $this; }

    public function getAdresse2(): ?string { return $this->adresse2; }

    public function setAdresse2(?string $adresse2): static { $this->adresse2 = $adresse2; return $this; }

    public function getCodePostal(): ?string { return $this->codePostal; }

    public function setCodePostal(?string $codePostal): static { $this->codePostal = $codePostal; return $this; }

    public function getVille(): ?string { return $this->ville; }

    public function setVille(?string $ville): static { $this->ville = $ville; return $this; }

    public function isReinstruction(): bool { return $this->reinstruction; }

    public function setReinstruction(bool $reinstruction): static { $this->reinstruction = $reinstruction; return $this; }

    public function getCommune(): ?Commune { return $this->commune; }

    public function setCommune(?Commune $commune): static { $this->commune = $commune; return $this; }

    public function getCanton(): ?Canton { return $this->canton; }

    public function setCanton(?Canton $canton): static { $this->canton = $canton; return $this; }

    public function getCirconscription(): ?Circonscription { return $this->circonscription; }

    public function setCirconscription(?Circonscription $circonscription): static { $this->circonscription = $circonscription; return $this; }

    public function getSituationFamiliale(): ?SituationFamiliale { return $this->situationFamiliale; }

    public function setSituationFamiliale(?SituationFamiliale $situationFamiliale): static { $this->situationFamiliale = $situationFamiliale; return $this; }

    public function getRegimeMatrimonial(): ?string { return $this->regimeMatrimonial; }

    public function setRegimeMatrimonial(?string $regimeMatrimonial): static { $this->regimeMatrimonial = $regimeMatrimonial; return $this; }

    public function getDateMariage(): ?\DateTimeImmutable { return $this->dateMariage; }

    public function setDateMariage(?\DateTimeImmutable $dateMariage): static { $this->dateMariage = $dateMariage; return $this; }

    public function getLieuMariage(): ?string { return $this->lieuMariage; }

    public function setLieuMariage(?string $lieuMariage): static { $this->lieuMariage = $lieuMariage; return $this; }

    /** @return Collection<int, Individu> */
    public function getIndividus(): Collection { return $this->individus; }

    public function addIndividu(Individu $item): static
    {
        if (!$this->individus->contains($item)) {
            $this->individus->add($item);
            $item->setDossier($this);
        }

        return $this;
    }

    /** @return Collection<int, Demande> */
    public function getDemandes(): Collection { return $this->demandes; }

    public function addDemande(Demande $item): static
    {
        if (!$this->demandes->contains($item)) {
            $this->demandes->add($item);
            $item->setDossier($this);
        }

        return $this;
    }

    public function __toString(): string { return (string) $this->numero; }

    /** Intitulé du dossier : « Nom Prénom » du demandeur, suivi du prénom du conjoint s'il existe. */
    public function getLibelle(): string
    {
        $demandeur = null;
        $autres = [];
        foreach ($this->individus as $i) {
            if (null === $demandeur && 'DE' === $i->getType()) {
                $demandeur = $i;
            } else {
                $autres[] = $i;
            }
        }
        $demandeur ??= array_shift($autres);
        if (null === $demandeur) {
            return 'Dossier '.$this->numero;
        }
        $libelle = trim($demandeur->getNom().' '.$demandeur->getPrenom());
        if ($autres) {
            $conjoint = $autres[0];
            $libelle .= ' & '.($conjoint->getNom() === $demandeur->getNom() ? (string) $conjoint->getPrenom() : trim($conjoint->getNom().' '.$conjoint->getPrenom()));
        }

        return $libelle;
    }

    public function getInitiales(): string
    {
        $mots = preg_split('/[\s&]+/u', $this->getLibelle(), -1, PREG_SPLIT_NO_EMPTY);

        return mb_strtoupper(mb_substr($mots[0] ?? '?', 0, 1).mb_substr($mots[1] ?? '', 0, 1));
    }
}

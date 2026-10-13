<?php

namespace App\Entity;

use Doctrine\Common\Collections\ArrayCollection;
use Doctrine\Common\Collections\Collection;
use Doctrine\ORM\Mapping as ORM;
use Symfony\Component\Validator\Constraints as Assert;

#[ORM\Entity]
class Individu
{
    #[ORM\Id, ORM\GeneratedValue, ORM\Column]
    private ?int $id = null;

    #[ORM\ManyToOne(targetEntity: Dossier::class, inversedBy: 'individus')]
    #[ORM\JoinColumn(nullable: false, onDelete: 'CASCADE')]
    private Dossier $dossier;

    #[ORM\Column(length: 2, nullable: false)]
    private string $type = '';

    #[ORM\Column(length: 50, nullable: false)]
    private string $nom = '';

    #[ORM\Column(length: 50, nullable: true)]
    private ?string $prenom = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateNaissance = null;

    #[ORM\Column(length: 100, nullable: true)]
    private ?string $lieuNaissance = null;

    #[ORM\Column(length: 50, nullable: true)]
    private ?string $nationalite = null;

    #[ORM\ManyToOne(targetEntity: Profession::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?Profession $profession = null;

    #[ORM\ManyToOne(targetEntity: RegimeAssuranceMaladie::class)]
    #[ORM\JoinColumn(nullable: true)]
    private ?RegimeAssuranceMaladie $regimeAssuranceMaladie = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateDeces = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateDivorce = null;

    #[ORM\Column(type: 'date_immutable', nullable: true)]
    private ?\DateTimeImmutable $dateVeuvage = null;

    #[ORM\Column(length: 150, nullable: true)]
    private ?string $email = null;

    #[ORM\Column(length: 150, nullable: true)]
    private ?string $adresse1 = null;

    #[ORM\Column(length: 150, nullable: true)]
    private ?string $adresse2 = null;

    #[ORM\Column(length: 10, nullable: true)]
    private ?string $codePostal = null;

    #[ORM\Column(length: 100, nullable: true)]
    private ?string $ville = null;

    public function getId(): ?int { return $this->id; }

    public function getDossier(): Dossier { return $this->dossier; }

    public function setDossier(Dossier $dossier): static { $this->dossier = $dossier; return $this; }

    public function getType(): string { return $this->type; }

    public function setType(string $type): static { $this->type = $type; return $this; }

    public function getNom(): string { return $this->nom; }

    public function setNom(string $nom): static { $this->nom = $nom; return $this; }

    public function getPrenom(): ?string { return $this->prenom; }

    public function setPrenom(?string $prenom): static { $this->prenom = $prenom; return $this; }

    public function getDateNaissance(): ?\DateTimeImmutable { return $this->dateNaissance; }

    public function setDateNaissance(?\DateTimeImmutable $dateNaissance): static { $this->dateNaissance = $dateNaissance; return $this; }

    public function getLieuNaissance(): ?string { return $this->lieuNaissance; }

    public function setLieuNaissance(?string $lieuNaissance): static { $this->lieuNaissance = $lieuNaissance; return $this; }

    public function getNationalite(): ?string { return $this->nationalite; }

    public function setNationalite(?string $nationalite): static { $this->nationalite = $nationalite; return $this; }

    public function getProfession(): ?Profession { return $this->profession; }

    public function setProfession(?Profession $profession): static { $this->profession = $profession; return $this; }

    public function getRegimeAssuranceMaladie(): ?RegimeAssuranceMaladie { return $this->regimeAssuranceMaladie; }

    public function setRegimeAssuranceMaladie(?RegimeAssuranceMaladie $regimeAssuranceMaladie): static { $this->regimeAssuranceMaladie = $regimeAssuranceMaladie; return $this; }

    public function getDateDeces(): ?\DateTimeImmutable { return $this->dateDeces; }

    public function setDateDeces(?\DateTimeImmutable $dateDeces): static { $this->dateDeces = $dateDeces; return $this; }

    public function getDateDivorce(): ?\DateTimeImmutable { return $this->dateDivorce; }

    public function setDateDivorce(?\DateTimeImmutable $dateDivorce): static { $this->dateDivorce = $dateDivorce; return $this; }

    public function getDateVeuvage(): ?\DateTimeImmutable { return $this->dateVeuvage; }

    public function setDateVeuvage(?\DateTimeImmutable $dateVeuvage): static { $this->dateVeuvage = $dateVeuvage; return $this; }

    public function getEmail(): ?string { return $this->email; }

    public function setEmail(?string $email): static { $this->email = $email; return $this; }

    public function getAdresse1(): ?string { return $this->adresse1; }

    public function setAdresse1(?string $adresse1): static { $this->adresse1 = $adresse1; return $this; }

    public function getAdresse2(): ?string { return $this->adresse2; }

    public function setAdresse2(?string $adresse2): static { $this->adresse2 = $adresse2; return $this; }

    public function getCodePostal(): ?string { return $this->codePostal; }

    public function setCodePostal(?string $codePostal): static { $this->codePostal = $codePostal; return $this; }

    public function getVille(): ?string { return $this->ville; }

    public function setVille(?string $ville): static { $this->ville = $ville; return $this; }
}

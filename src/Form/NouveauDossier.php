<?php

namespace App\Form;

use App\Entity\Commune;
use App\Entity\NatureDemande;
use App\Entity\OrigineEthnique;
use App\Entity\Profession;
use App\Entity\SituationFamiliale;
use App\Entity\TypeDemande;
use Symfony\Component\Validator\Constraints as Assert;

/** Données saisies dans l'assistant de création d'un dossier (demandeur, foyer, projet). */
class NouveauDossier
{
    #[Assert\NotBlank(message: 'Indiquez le numéro du dossier.')]
    public ?string $numero = null;

    #[Assert\NotBlank(message: 'Indiquez le nom.')]
    public ?string $nom = null;
    public ?string $prenom = null;
    public ?\DateTimeImmutable $dateNaissance = null;
    public ?Profession $profession = null;
    public ?string $adresse1 = null;
    public ?string $codePostal = null;
    public ?string $ville = null;
    public ?Commune $commune = null;
    public ?string $telephone = null;

    public ?SituationFamiliale $situationFamiliale = null;
    public ?\DateTimeImmutable $dateMariage = null;
    public ?string $lieuMariage = null;
    public bool $conjoint = false;
    public ?string $conjointNom = null;
    public ?string $conjointPrenom = null;
    public ?\DateTimeImmutable $conjointDateNaissance = null;

    #[Assert\NotNull(message: 'Indiquez la date de la demande.')]
    public ?\DateTimeImmutable $dateDemande = null;
    public ?NatureDemande $nature = null;
    public ?TypeDemande $type = null;
    #[Assert\PositiveOrZero]
    public ?int $nombreEnfants = 1;
    public ?OrigineEthnique $origineEthnique = null;
    public bool $fratrie = false;
}

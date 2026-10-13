<?php

namespace App\Form;

use App\Entity\Commune;
use App\Entity\NatureDemande;
use App\Entity\OrigineEthnique;
use App\Entity\Profession;
use App\Entity\SituationFamiliale;
use App\Entity\TypeDemande;
use Symfony\Bridge\Doctrine\Form\Type\EntityType;
use Symfony\Component\Form\AbstractType;
use Symfony\Component\Form\Extension\Core\Type as T;
use Symfony\Component\Form\FormBuilderInterface;
use Symfony\Component\OptionsResolver\OptionsResolver;

class NouveauDossierType extends AbstractType
{
    public function buildForm(FormBuilderInterface $builder, array $options): void
    {
        $date = ['widget' => 'single_text', 'input' => 'datetime_immutable', 'required' => false];
        $entity = static fn (string $class, string $label, bool $required = false): array => ['class' => $class, 'required' => $required, 'placeholder' => '—', 'label' => $label];

        $builder
            ->add('numero', T\TextType::class, ['label' => 'Numéro de dossier'])
            ->add('nom', T\TextType::class, ['label' => 'Nom'])
            ->add('prenom', T\TextType::class, ['label' => 'Prénom', 'required' => false])
            ->add('dateNaissance', T\DateType::class, $date + ['label' => 'Date de naissance'])
            ->add('profession', EntityType::class, $entity(Profession::class, 'Profession'))
            ->add('adresse1', T\TextType::class, ['label' => 'Adresse', 'required' => false])
            ->add('codePostal', T\TextType::class, ['label' => 'Code postal', 'required' => false])
            ->add('ville', T\TextType::class, ['label' => 'Ville', 'required' => false])
            ->add('commune', EntityType::class, $entity(Commune::class, 'Commune') + ['help' => 'Le canton et la circonscription sont déduits de la commune.'])
            ->add('telephone', T\TextType::class, ['label' => 'Téléphone', 'required' => false])
            ->add('situationFamiliale', EntityType::class, $entity(SituationFamiliale::class, 'Situation familiale'))
            ->add('dateMariage', T\DateType::class, $date + ['label' => 'Date du mariage'])
            ->add('lieuMariage', T\TextType::class, ['label' => 'Lieu du mariage', 'required' => false])
            ->add('conjoint', T\CheckboxType::class, ['label' => 'Ajouter le conjoint ou la conjointe', 'required' => false])
            ->add('conjointNom', T\TextType::class, ['label' => 'Nom', 'required' => false])
            ->add('conjointPrenom', T\TextType::class, ['label' => 'Prénom', 'required' => false])
            ->add('conjointDateNaissance', T\DateType::class, $date + ['label' => 'Date de naissance'])
            ->add('dateDemande', T\DateType::class, $date + ['label' => 'Date de la demande'])
            ->add('nature', EntityType::class, $entity(NatureDemande::class, 'Nature de la demande'))
            ->add('type', EntityType::class, $entity(TypeDemande::class, 'Type de demande'))
            ->add('nombreEnfants', T\IntegerType::class, ['label' => 'Nombre d\'enfants souhaités', 'required' => false])
            ->add('origineEthnique', EntityType::class, $entity(OrigineEthnique::class, 'Origine ethnique souhaitée'))
            ->add('fratrie', T\CheckboxType::class, ['label' => 'Fratrie acceptée', 'required' => false]);
    }

    public function configureOptions(OptionsResolver $resolver): void
    {
        $resolver->setDefaults(['data_class' => NouveauDossier::class]);
    }
}

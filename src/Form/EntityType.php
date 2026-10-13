<?php

namespace App\Form;

use Doctrine\ORM\EntityManagerInterface;
use Doctrine\ORM\Mapping\ClassMetadata;
use Symfony\Bridge\Doctrine\Form\Type\EntityType as DoctrineEntityType;
use Symfony\Component\Form\AbstractType;
use Symfony\Component\Form\Extension\Core\Type as T;
use Symfony\Component\Form\FormBuilderInterface;
use Symfony\Component\OptionsResolver\OptionsResolver;

/** Formulaire générique construit à partir des métadonnées Doctrine de l'entité. */
class EntityType extends AbstractType
{
    private const CHOIX = [
        'sexe' => ['Indéterminé' => 1, 'Masculin' => 2, 'Féminin' => 3],
        'deliberation' => ['Indéterminée' => 0, 'Refus' => 1, 'Retrait' => 2, 'Accord' => 3],
        'cadre' => ['Demande initiale' => 1, 'Seconde évaluation' => 2, 'Reprise' => 3, 'Recours administratif' => 4, 'Surveillance de placement' => 5],
        'type' => ['Psychologue' => 'psychologue', 'Travailleur social' => 'assistante_sociale'],
        'provenance' => ['Commission initiale' => -2, 'Recours administratif' => -1, 'Recours contentieux' => 0],
    ];

    public function __construct(private readonly EntityManagerInterface $em)
    {
    }

    public function buildForm(FormBuilderInterface $builder, array $options): void
    {
        /** @var ClassMetadata $meta */
        $meta = $this->em->getClassMetadata($options['data_class']);

        foreach ($meta->getFieldNames() as $name) {
            if ($meta->isIdentifier($name) && !($options['id_editable'] && 'id' === $name)) {
                continue;
            }
            $mapping = $meta->getFieldMapping($name);
            $type = $mapping->type;
            $required = !$mapping->nullable && 'boolean' !== $type;
            $opts = ['required' => $required, 'label' => self::label($name)];
            if (isset(self::CHOIX[$name]) && 'Individu' !== $meta->getReflectionClass()->getShortName()) {
                $builder->add($name, T\ChoiceType::class, $opts + ['choices' => self::CHOIX[$name]]);
                continue;
            }
            $builder->add($name, match ($type) {
                'boolean' => T\CheckboxType::class,
                'text' => T\TextareaType::class,
                'integer' => T\IntegerType::class,
                'date_immutable' => T\DateType::class,
                default => T\TextType::class,
            }, $opts + ('date_immutable' === $type ? ['widget' => 'single_text', 'input' => 'datetime_immutable'] : []));
        }

        foreach ($meta->getAssociationNames() as $name) {
            if ($name === $options['exclude'] || !$meta->isSingleValuedAssociation($name)) {
                continue;
            }
            $builder->add($name, DoctrineEntityType::class, [
                'class' => $meta->getAssociationTargetClass($name),
                'required' => false,
                'placeholder' => '—',
                'label' => self::label($name),
            ]);
        }
    }

    public function configureOptions(OptionsResolver $resolver): void
    {
        $resolver->setDefaults(['exclude' => null, 'id_editable' => false]);
        $resolver->setRequired('data_class');
    }

    private static function label(string $name): string
    {
        return ucfirst(trim(strtolower((string) preg_replace('/(?<!^)[A-Z]/', ' $0', $name))));
    }
}

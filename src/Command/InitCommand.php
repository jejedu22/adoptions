<?php

namespace App\Command;

use App\Entity\Parametre;
use App\Entity\TypeRecours;
use Doctrine\ORM\EntityManagerInterface;
use Symfony\Component\Console\Attribute\AsCommand;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;
use Symfony\Component\Console\Style\SymfonyStyle;

#[AsCommand(name: 'app:init', description: 'Crée les paramètres de l\'échéancier et les types de recours par défaut')]
class InitCommand extends Command
{
    private const PARAMETRES = [
        1 => ['Demande → réception du dossier', 1],
        2 => ['Réception du dossier → réponse du travailleur social', 3],
        3 => ['Réception du dossier → réponse du psychologue', 3],
        4 => ['Réception du dossier → décision', 9],
        5 => ['Demande de renouvellement → réponse du travailleur social', 3],
        6 => ['Deux surveillances avant jugement', 6],
        7 => ['Recours administratif → réponse du travailleur social', 2],
        8 => ['Recours administratif → réponse du psychologue', 2],
    ];

    private const RECOURS = [
        'Tribunal administratif' => false,
        'Cour administrative d\'appel' => false,
        'Conseil d\'État' => true,
    ];

    public function __construct(private readonly EntityManagerInterface $em)
    {
        parent::__construct();
    }

    protected function execute(InputInterface $input, OutputInterface $output): int
    {
        $io = new SymfonyStyle($input, $output);
        foreach (self::PARAMETRES as $id => [$libelle, $mois]) {
            if (null === $this->em->find(Parametre::class, $id)) {
                $this->em->persist((new Parametre())->setId($id)->setLibelle($libelle)->setNombreMois($mois));
            }
        }
        foreach (self::RECOURS as $libelle => $conseilEtat) {
            if (null === $this->em->getRepository(TypeRecours::class)->findOneBy(['libelle' => $libelle])) {
                $this->em->persist((new TypeRecours())->setLibelle($libelle)->setConseilEtat($conseilEtat));
            }
        }
        $this->em->flush();
        $io->success('Données par défaut créées.');

        return Command::SUCCESS;
    }
}

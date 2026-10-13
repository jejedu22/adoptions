<?php

namespace App\Command;

use App\Import\LegacyImporter;
use App\Import\LegacyMapping;
use Doctrine\DBAL\DriverManager;
use Doctrine\DBAL\Tools\DsnParser;
use Symfony\Component\Console\Attribute\AsCommand;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Output\OutputInterface;
use Symfony\Component\Console\Style\SymfonyStyle;

#[AsCommand(name: 'app:import-legacy', description: 'Importe les données de l\'ancienne base Access (copiée dans MySQL) dans la base de l\'application')]
class ImportLegacyCommand extends Command
{
    public function __construct(private readonly LegacyImporter $importer)
    {
        parent::__construct();
    }

    protected function configure(): void
    {
        $this
            ->addOption('source', 's', InputOption::VALUE_REQUIRED, 'DSN de la base source, ex. mysql://user:mdp@127.0.0.1/adoptions_legacy (sinon variable LEGACY_DATABASE_URL)')
            ->addOption('check', null, InputOption::VALUE_NONE, 'Vérifie seulement les tables et colonnes de la source, sans rien écrire')
            ->addOption('dry-run', null, InputOption::VALUE_NONE, 'Exécute l\'import puis annule tout (aucune écriture)')
            ->addOption('strict', null, InputOption::VALUE_NONE, 'Échoue si une table ou une colonne attendue est absente de la source');
    }

    protected function execute(InputInterface $input, OutputInterface $output): int
    {
        $io = new SymfonyStyle($input, $output);
        $dsn = $input->getOption('source') ?: ($_SERVER['LEGACY_DATABASE_URL'] ?? $_ENV['LEGACY_DATABASE_URL'] ?? null);
        if (!$dsn) {
            $io->error('Indiquez la base source avec --source ou la variable LEGACY_DATABASE_URL.');

            return Command::INVALID;
        }

        $source = DriverManager::getConnection((new DsnParser(['mysql' => 'pdo_mysql', 'mariadb' => 'pdo_mysql', 'sqlite' => 'pdo_sqlite', 'pgsql' => 'pdo_pgsql']))->parse($dsn));
        $check = (bool) $input->getOption('check');
        $dry = (bool) $input->getOption('dry-run');

        try {
            $reports = $this->importer->import($source, LegacyMapping::steps(), $check, (bool) $input->getOption('strict'), $dry);
        } catch (\Throwable $e) {
            $io->error($e->getMessage());

            return Command::FAILURE;
        }

        $rows = [];
        $warnings = [];
        foreach ($reports as $r) {
            if ($r->tableAbsente) {
                $rows[] = [$r->step->name, $r->step->table, 'absente', '', '', '', '', ''];
                $warnings[] = sprintf('Table %s absente : « %s » n\'est pas importé.', $r->step->table, $r->step->name);
                continue;
            }
            $rows[] = $check
                ? [$r->step->name, $r->table, $r->nbLignes, '', '', '', '', count($r->colonnesAbsentes) ?: '']
                : [$r->step->name, $r->table, $r->lues, $r->crees, $r->misAJour, $r->ignorees, $r->refsIntrouvables, count($r->colonnesAbsentes) ?: ''];
            if ($r->colonnesAbsentes) {
                $warnings[] = sprintf('%s : colonnes absentes de la source (ignorées) : %s', $r->table, implode(', ', $r->colonnesAbsentes));
            }
        }
        $io->table(['Étape', 'Table', $check ? 'Lignes' : 'Lues', 'Créées', 'Mises à jour', 'Ignorées', 'Réf. introuvables', 'Col. absentes'], $rows);
        foreach ($warnings as $w) {
            $io->warning($w);
        }

        if ($check) {
            $io->note('Vérification seule : rien n\'a été importé. Ajustez src/Import/LegacyMapping.php si des colonnes sont absentes.');
        } elseif ($dry) {
            $io->note('Mode --dry-run : toutes les écritures ont été annulées.');
        } else {
            $io->success('Import terminé. Il peut être relancé sans créer de doublons.');
        }

        return Command::SUCCESS;
    }
}

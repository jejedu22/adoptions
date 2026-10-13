<?php

declare(strict_types=1);

namespace DoctrineMigrations;

use Doctrine\DBAL\Schema\Schema;
use Doctrine\Migrations\AbstractMigration;

/**
 * Auto-generated Migration: Please modify to your needs!
 */
final class Version20261013140908 extends AbstractMigration
{
    public function getDescription(): string
    {
        return '';
    }

    public function up(Schema $schema): void
    {
        // this up() migration is auto-generated, please modify it to your needs
        $this->addSql('CREATE TABLE legacy_map (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, entity VARCHAR(150) NOT NULL, legacy_key VARCHAR(191) NOT NULL, new_id INTEGER NOT NULL)');
        $this->addSql('CREATE UNIQUE INDEX UNIQ_4C6A3587E28446855573495 ON legacy_map (entity, legacy_key)');
        $this->addSql('CREATE TEMPORARY TABLE __temp__parametre AS SELECT id, libelle, nombre_mois FROM parametre');
        $this->addSql('DROP TABLE parametre');
        $this->addSql('CREATE TABLE parametre (id INTEGER NOT NULL, libelle VARCHAR(150) NOT NULL, nombre_mois INTEGER NOT NULL, PRIMARY KEY (id))');
        $this->addSql('INSERT INTO parametre (id, libelle, nombre_mois) SELECT id, libelle, nombre_mois FROM __temp__parametre');
        $this->addSql('DROP TABLE __temp__parametre');
    }

    public function down(Schema $schema): void
    {
        // this down() migration is auto-generated, please modify it to your needs
        $this->addSql('DROP TABLE legacy_map');
        $this->addSql('CREATE TEMPORARY TABLE __temp__parametre AS SELECT id, libelle, nombre_mois FROM parametre');
        $this->addSql('DROP TABLE parametre');
        $this->addSql('CREATE TABLE parametre (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, libelle VARCHAR(150) NOT NULL, nombre_mois INTEGER NOT NULL)');
        $this->addSql('INSERT INTO parametre (id, libelle, nombre_mois) SELECT id, libelle, nombre_mois FROM __temp__parametre');
        $this->addSql('DROP TABLE __temp__parametre');
    }
}

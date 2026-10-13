<?php

declare(strict_types=1);

namespace DoctrineMigrations;

use Doctrine\DBAL\Schema\Schema;
use Doctrine\Migrations\AbstractMigration;

/**
 * Auto-generated Migration: Please modify to your needs!
 */
final class Version20261013140856 extends AbstractMigration
{
    public function getDescription(): string
    {
        return '';
    }

    public function up(Schema $schema): void
    {
        // this up() migration is auto-generated, please modify it to your needs
        $this->addSql('CREATE TABLE canton (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, nom VARCHAR(150) NOT NULL)');
        $this->addSql('CREATE TABLE circonscription (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, nom VARCHAR(150) NOT NULL, adresse VARCHAR(200) DEFAULT NULL, code_postal VARCHAR(10) DEFAULT NULL, ville VARCHAR(100) DEFAULT NULL, responsable VARCHAR(150) DEFAULT NULL)');
        $this->addSql('CREATE TABLE commune (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, nom VARCHAR(150) NOT NULL, code_postal VARCHAR(10) DEFAULT NULL, code_insee VARCHAR(10) DEFAULT NULL, canton_id INTEGER DEFAULT NULL, circonscription_id INTEGER DEFAULT NULL, CONSTRAINT FK_E2E2D1EE8D070D0B FOREIGN KEY (canton_id) REFERENCES canton (id) NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_E2E2D1EE755DBAE FOREIGN KEY (circonscription_id) REFERENCES circonscription (id) NOT DEFERRABLE INITIALLY IMMEDIATE)');
        $this->addSql('CREATE INDEX IDX_E2E2D1EE8D070D0B ON commune (canton_id)');
        $this->addSql('CREATE INDEX IDX_E2E2D1EE755DBAE ON commune (circonscription_id)');
        $this->addSql('CREATE TABLE decision (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, provenance INTEGER NOT NULL, commission_initiale DATE DEFAULT NULL, date_commission DATE DEFAULT NULL, date_signature DATE DEFAULT NULL, date_depot DATE DEFAULT NULL, date_decision_prevue DATE DEFAULT NULL, deliberation INTEGER NOT NULL, remarques CLOB DEFAULT NULL, demande_id INTEGER NOT NULL, type_recours_id INTEGER DEFAULT NULL, CONSTRAINT FK_84ACBE4880E95E18 FOREIGN KEY (demande_id) REFERENCES demande (id) ON DELETE CASCADE NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_84ACBE4812C88C4 FOREIGN KEY (type_recours_id) REFERENCES type_recours (id) NOT DEFERRABLE INITIALLY IMMEDIATE)');
        $this->addSql('CREATE INDEX IDX_84ACBE4880E95E18 ON decision (demande_id)');
        $this->addSql('CREATE INDEX IDX_84ACBE4812C88C4 ON decision (type_recours_id)');
        $this->addSql('CREATE TABLE demande (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, numero INTEGER NOT NULL, date_demande DATE DEFAULT NULL, nombre_enfants INTEGER DEFAULT NULL, fratrie BOOLEAN NOT NULL, assistante_maternelle BOOLEAN NOT NULL, remarques CLOB DEFAULT NULL, date_sortie DATE DEFAULT NULL, motif_sortie VARCHAR(255) DEFAULT NULL, date_suspension DATE DEFAULT NULL, motif_suspension VARCHAR(255) DEFAULT NULL, date_agrement DATE DEFAULT NULL, provenance_agrement INTEGER DEFAULT NULL, date_retrait DATE DEFAULT NULL, date_prevue_decision DATE DEFAULT NULL, date_reception_dossier DATE DEFAULT NULL, a_actualiser BOOLEAN NOT NULL, dossier_id INTEGER NOT NULL, origine_ethnique_id INTEGER DEFAULT NULL, nature_id INTEGER DEFAULT NULL, type_id INTEGER DEFAULT NULL, CONSTRAINT FK_2694D7A5611C0C56 FOREIGN KEY (dossier_id) REFERENCES dossier (id) ON DELETE CASCADE NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_2694D7A540149531 FOREIGN KEY (origine_ethnique_id) REFERENCES origine_ethnique (id) NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_2694D7A53BCB2E4B FOREIGN KEY (nature_id) REFERENCES nature_demande (id) NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_2694D7A5C54C8C93 FOREIGN KEY (type_id) REFERENCES type_demande (id) NOT DEFERRABLE INITIALLY IMMEDIATE)');
        $this->addSql('CREATE INDEX IDX_2694D7A5611C0C56 ON demande (dossier_id)');
        $this->addSql('CREATE INDEX IDX_2694D7A540149531 ON demande (origine_ethnique_id)');
        $this->addSql('CREATE INDEX IDX_2694D7A53BCB2E4B ON demande (nature_id)');
        $this->addSql('CREATE INDEX IDX_2694D7A5C54C8C93 ON demande (type_id)');
        $this->addSql('CREATE TABLE dossier (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, numero VARCHAR(20) NOT NULL, departement VARCHAR(3) DEFAULT NULL, telephone VARCHAR(30) DEFAULT NULL, adresse1 VARCHAR(150) DEFAULT NULL, adresse2 VARCHAR(150) DEFAULT NULL, code_postal VARCHAR(10) DEFAULT NULL, ville VARCHAR(100) DEFAULT NULL, reinstruction BOOLEAN NOT NULL, regime_matrimonial VARCHAR(150) DEFAULT NULL, date_mariage DATE DEFAULT NULL, lieu_mariage VARCHAR(100) DEFAULT NULL, commune_id INTEGER DEFAULT NULL, canton_id INTEGER DEFAULT NULL, circonscription_id INTEGER DEFAULT NULL, situation_familiale_id INTEGER DEFAULT NULL, CONSTRAINT FK_3D48E037131A4F72 FOREIGN KEY (commune_id) REFERENCES commune (id) NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_3D48E0378D070D0B FOREIGN KEY (canton_id) REFERENCES canton (id) NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_3D48E037755DBAE FOREIGN KEY (circonscription_id) REFERENCES circonscription (id) NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_3D48E037F11CEA43 FOREIGN KEY (situation_familiale_id) REFERENCES situation_familiale (id) NOT DEFERRABLE INITIALLY IMMEDIATE)');
        $this->addSql('CREATE INDEX IDX_3D48E037131A4F72 ON dossier (commune_id)');
        $this->addSql('CREATE INDEX IDX_3D48E0378D070D0B ON dossier (canton_id)');
        $this->addSql('CREATE INDEX IDX_3D48E037755DBAE ON dossier (circonscription_id)');
        $this->addSql('CREATE INDEX IDX_3D48E037F11CEA43 ON dossier (situation_familiale_id)');
        $this->addSql('CREATE TABLE enfant_demande (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, sexe INTEGER NOT NULL, age_max INTEGER DEFAULT NULL, handicap BOOLEAN NOT NULL, demande_id INTEGER NOT NULL, origine_ethnique_id INTEGER DEFAULT NULL, CONSTRAINT FK_2E7A81D680E95E18 FOREIGN KEY (demande_id) REFERENCES demande (id) ON DELETE CASCADE NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_2E7A81D640149531 FOREIGN KEY (origine_ethnique_id) REFERENCES origine_ethnique (id) NOT DEFERRABLE INITIALLY IMMEDIATE)');
        $this->addSql('CREATE INDEX IDX_2E7A81D680E95E18 ON enfant_demande (demande_id)');
        $this->addSql('CREATE INDEX IDX_2E7A81D640149531 ON enfant_demande (origine_ethnique_id)');
        $this->addSql('CREATE TABLE famille_document (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, libelle VARCHAR(150) NOT NULL)');
        $this->addSql('CREATE TABLE historique_intervenant (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, cadre INTEGER NOT NULL, date_demande DATE DEFAULT NULL, date_prevue DATE DEFAULT NULL, date_recu DATE DEFAULT NULL, demande_id INTEGER NOT NULL, intervenant_id INTEGER NOT NULL, CONSTRAINT FK_AA7E1A3A80E95E18 FOREIGN KEY (demande_id) REFERENCES demande (id) ON DELETE CASCADE NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_AA7E1A3AAB9A1716 FOREIGN KEY (intervenant_id) REFERENCES intervenant (id) NOT DEFERRABLE INITIALLY IMMEDIATE)');
        $this->addSql('CREATE INDEX IDX_AA7E1A3A80E95E18 ON historique_intervenant (demande_id)');
        $this->addSql('CREATE INDEX IDX_AA7E1A3AAB9A1716 ON historique_intervenant (intervenant_id)');
        $this->addSql('CREATE TABLE individu (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, type VARCHAR(2) NOT NULL, nom VARCHAR(50) NOT NULL, prenom VARCHAR(50) DEFAULT NULL, date_naissance DATE DEFAULT NULL, lieu_naissance VARCHAR(100) DEFAULT NULL, nationalite VARCHAR(50) DEFAULT NULL, date_deces DATE DEFAULT NULL, date_divorce DATE DEFAULT NULL, date_veuvage DATE DEFAULT NULL, email VARCHAR(150) DEFAULT NULL, adresse1 VARCHAR(150) DEFAULT NULL, adresse2 VARCHAR(150) DEFAULT NULL, code_postal VARCHAR(10) DEFAULT NULL, ville VARCHAR(100) DEFAULT NULL, dossier_id INTEGER NOT NULL, profession_id INTEGER DEFAULT NULL, regime_assurance_maladie_id INTEGER DEFAULT NULL, CONSTRAINT FK_5EE42FCE611C0C56 FOREIGN KEY (dossier_id) REFERENCES dossier (id) ON DELETE CASCADE NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_5EE42FCEFDEF8996 FOREIGN KEY (profession_id) REFERENCES profession (id) NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_5EE42FCE50A550A7 FOREIGN KEY (regime_assurance_maladie_id) REFERENCES regime_assurance_maladie (id) NOT DEFERRABLE INITIALLY IMMEDIATE)');
        $this->addSql('CREATE INDEX IDX_5EE42FCE611C0C56 ON individu (dossier_id)');
        $this->addSql('CREATE INDEX IDX_5EE42FCEFDEF8996 ON individu (profession_id)');
        $this->addSql('CREATE INDEX IDX_5EE42FCE50A550A7 ON individu (regime_assurance_maladie_id)');
        $this->addSql('CREATE TABLE institution (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, libelle VARCHAR(150) NOT NULL)');
        $this->addSql('CREATE TABLE intervenant (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, nom VARCHAR(100) NOT NULL, titre VARCHAR(20) DEFAULT NULL, telephone VARCHAR(30) DEFAULT NULL, actif BOOLEAN NOT NULL, type VARCHAR(30) NOT NULL, qualite_id INTEGER DEFAULT NULL, institution_id INTEGER DEFAULT NULL, permanence_id INTEGER DEFAULT NULL, CONSTRAINT FK_73D0145CA6338570 FOREIGN KEY (qualite_id) REFERENCES qualite (id) NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_73D0145C10405986 FOREIGN KEY (institution_id) REFERENCES institution (id) NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_73D0145CA9457964 FOREIGN KEY (permanence_id) REFERENCES permanence (id) NOT DEFERRABLE INITIALLY IMMEDIATE)');
        $this->addSql('CREATE INDEX IDX_73D0145CA6338570 ON intervenant (qualite_id)');
        $this->addSql('CREATE INDEX IDX_73D0145C10405986 ON intervenant (institution_id)');
        $this->addSql('CREATE INDEX IDX_73D0145CA9457964 ON intervenant (permanence_id)');
        $this->addSql('CREATE TABLE legacy_map (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, entity VARCHAR(150) NOT NULL, legacy_key VARCHAR(191) NOT NULL, new_id INTEGER NOT NULL)');
        $this->addSql('CREATE UNIQUE INDEX UNIQ_4C6A3587E28446855573495 ON legacy_map (entity, legacy_key)');
        $this->addSql('CREATE TABLE nature_demande (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, libelle VARCHAR(150) NOT NULL)');
        $this->addSql('CREATE TABLE oeuvre (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, designation VARCHAR(150) NOT NULL, responsable VARCHAR(150) DEFAULT NULL, adresse VARCHAR(200) DEFAULT NULL, code_postal VARCHAR(10) DEFAULT NULL, ville VARCHAR(100) DEFAULT NULL, telephone VARCHAR(30) DEFAULT NULL, email VARCHAR(150) DEFAULT NULL, remarques CLOB DEFAULT NULL)');
        $this->addSql('CREATE TABLE origine_ethnique (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, libelle VARCHAR(150) NOT NULL)');
        $this->addSql('CREATE TABLE parametre (id INTEGER NOT NULL, libelle VARCHAR(150) NOT NULL, nombre_mois INTEGER NOT NULL, PRIMARY KEY (id))');
        $this->addSql('CREATE TABLE permanence (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, lieu VARCHAR(150) NOT NULL, ville VARCHAR(100) DEFAULT NULL)');
        $this->addSql('CREATE TABLE placement (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, nom_enfant VARCHAR(100) DEFAULT NULL, prenom_enfant VARCHAR(100) DEFAULT NULL, date_naissance DATE DEFAULT NULL, sexe INTEGER NOT NULL, date_placement DATE DEFAULT NULL, date_jugement DATE DEFAULT NULL, pupille BOOLEAN NOT NULL, remarques CLOB DEFAULT NULL, demande_id INTEGER NOT NULL, origine_ethnique_id INTEGER DEFAULT NULL, oeuvre_id INTEGER DEFAULT NULL, tribunal_id INTEGER DEFAULT NULL, CONSTRAINT FK_48DB750E80E95E18 FOREIGN KEY (demande_id) REFERENCES demande (id) ON DELETE CASCADE NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_48DB750E40149531 FOREIGN KEY (origine_ethnique_id) REFERENCES origine_ethnique (id) NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_48DB750E88194DE8 FOREIGN KEY (oeuvre_id) REFERENCES oeuvre (id) NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_48DB750E75C2CEC3 FOREIGN KEY (tribunal_id) REFERENCES tribunal (id) NOT DEFERRABLE INITIALLY IMMEDIATE)');
        $this->addSql('CREATE INDEX IDX_48DB750E80E95E18 ON placement (demande_id)');
        $this->addSql('CREATE INDEX IDX_48DB750E40149531 ON placement (origine_ethnique_id)');
        $this->addSql('CREATE INDEX IDX_48DB750E88194DE8 ON placement (oeuvre_id)');
        $this->addSql('CREATE INDEX IDX_48DB750E75C2CEC3 ON placement (tribunal_id)');
        $this->addSql('CREATE TABLE profession (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, libelle VARCHAR(150) NOT NULL)');
        $this->addSql('CREATE TABLE qualite (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, libelle VARCHAR(150) NOT NULL)');
        $this->addSql('CREATE TABLE regime_assurance_maladie (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, libelle VARCHAR(150) NOT NULL)');
        $this->addSql('CREATE TABLE situation_familiale (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, libelle VARCHAR(150) NOT NULL)');
        $this->addSql('CREATE TABLE source_donnees (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, libelle VARCHAR(150) NOT NULL)');
        $this->addSql('CREATE TABLE surveillance (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, date_prevue DATE DEFAULT NULL, date_realisee DATE DEFAULT NULL, date_envoi_tgi DATE DEFAULT NULL, remarques CLOB DEFAULT NULL, placement_id INTEGER NOT NULL, intervenant_id INTEGER DEFAULT NULL, CONSTRAINT FK_C17BAD5B2F966E9D FOREIGN KEY (placement_id) REFERENCES placement (id) ON DELETE CASCADE NOT DEFERRABLE INITIALLY IMMEDIATE, CONSTRAINT FK_C17BAD5BAB9A1716 FOREIGN KEY (intervenant_id) REFERENCES intervenant (id) NOT DEFERRABLE INITIALLY IMMEDIATE)');
        $this->addSql('CREATE INDEX IDX_C17BAD5B2F966E9D ON surveillance (placement_id)');
        $this->addSql('CREATE INDEX IDX_C17BAD5BAB9A1716 ON surveillance (intervenant_id)');
        $this->addSql('CREATE TABLE tribunal (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, libelle VARCHAR(150) NOT NULL, adresse1 VARCHAR(150) DEFAULT NULL, adresse2 VARCHAR(150) DEFAULT NULL, code_postal VARCHAR(10) DEFAULT NULL, ville VARCHAR(100) DEFAULT NULL, commentaires CLOB DEFAULT NULL)');
        $this->addSql('CREATE TABLE type_demande (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, libelle VARCHAR(150) NOT NULL)');
        $this->addSql('CREATE TABLE type_recours (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, libelle VARCHAR(150) NOT NULL, conseil_etat BOOLEAN NOT NULL)');
    }

    public function down(Schema $schema): void
    {
        // this down() migration is auto-generated, please modify it to your needs
        $this->addSql('DROP TABLE canton');
        $this->addSql('DROP TABLE circonscription');
        $this->addSql('DROP TABLE commune');
        $this->addSql('DROP TABLE decision');
        $this->addSql('DROP TABLE demande');
        $this->addSql('DROP TABLE dossier');
        $this->addSql('DROP TABLE enfant_demande');
        $this->addSql('DROP TABLE famille_document');
        $this->addSql('DROP TABLE historique_intervenant');
        $this->addSql('DROP TABLE individu');
        $this->addSql('DROP TABLE institution');
        $this->addSql('DROP TABLE intervenant');
        $this->addSql('DROP TABLE legacy_map');
        $this->addSql('DROP TABLE nature_demande');
        $this->addSql('DROP TABLE oeuvre');
        $this->addSql('DROP TABLE origine_ethnique');
        $this->addSql('DROP TABLE parametre');
        $this->addSql('DROP TABLE permanence');
        $this->addSql('DROP TABLE placement');
        $this->addSql('DROP TABLE profession');
        $this->addSql('DROP TABLE qualite');
        $this->addSql('DROP TABLE regime_assurance_maladie');
        $this->addSql('DROP TABLE situation_familiale');
        $this->addSql('DROP TABLE source_donnees');
        $this->addSql('DROP TABLE surveillance');
        $this->addSql('DROP TABLE tribunal');
        $this->addSql('DROP TABLE type_demande');
        $this->addSql('DROP TABLE type_recours');
    }
}

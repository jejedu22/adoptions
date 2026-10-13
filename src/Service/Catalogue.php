<?php

namespace App\Service;

use App\Entity as E;

/**
 * Catalogue des écrans de gestion : une entrée par entité (équivalent des menus et fiches de l'ancienne application).
 *
 * Clés : class, title, groupe (menu), colonnes (liste), parent (propriété vers le parent pour les entités filles).
 */
final class Catalogue
{
    /** @return array<string, array{class: class-string, title: string, groupe: string, colonnes: list<string>, parent?: string, parentSlug?: string}> */
    public static function all(): array
    {
        return [
            'regimes-assurance-maladie' => ['class' => E\RegimeAssuranceMaladie::class, 'title' => 'Régimes d\'assurance maladie', 'groupe' => 'Définir', 'colonnes' => ['libelle']],
            'situations-familiales' => ['class' => E\SituationFamiliale::class, 'title' => 'Situations familiales', 'groupe' => 'Définir', 'colonnes' => ['libelle']],
            'professions' => ['class' => E\Profession::class, 'title' => 'Professions', 'groupe' => 'Définir', 'colonnes' => ['libelle']],
            'natures-demandes' => ['class' => E\NatureDemande::class, 'title' => 'Natures des demandes/placements', 'groupe' => 'Définir', 'colonnes' => ['libelle']],
            'types-demandes' => ['class' => E\TypeDemande::class, 'title' => 'Types des demandes', 'groupe' => 'Définir', 'colonnes' => ['libelle']],
            'types-recours' => ['class' => E\TypeRecours::class, 'title' => 'Types des recours contentieux', 'groupe' => 'Définir', 'colonnes' => ['libelle', 'conseilEtat']],
            'origines-ethniques' => ['class' => E\OrigineEthnique::class, 'title' => 'Origines ethniques des enfants', 'groupe' => 'Définir', 'colonnes' => ['libelle']],
            'institutions' => ['class' => E\Institution::class, 'title' => 'Institutions des travailleurs sociaux', 'groupe' => 'Définir', 'colonnes' => ['libelle']],
            'qualites' => ['class' => E\Qualite::class, 'title' => 'Qualités des travailleurs sociaux', 'groupe' => 'Définir', 'colonnes' => ['libelle']],
            'permanences' => ['class' => E\Permanence::class, 'title' => 'Lieux de permanences', 'groupe' => 'Définir', 'colonnes' => ['lieu', 'ville']],
            'parametres' => ['class' => E\Parametre::class, 'title' => 'Paramètres de l\'échéancier', 'groupe' => 'Définir', 'colonnes' => ['id', 'libelle', 'nombreMois']],
            'intervenants' => ['class' => E\Intervenant::class, 'title' => 'Intervenants', 'groupe' => 'Préparer', 'colonnes' => ['nom', 'titre', 'type', 'qualite', 'actif']],
            'oeuvres' => ['class' => E\Oeuvre::class, 'title' => 'Œuvres de placement', 'groupe' => 'Préparer', 'colonnes' => ['designation', 'responsable', 'ville']],
            'tribunaux' => ['class' => E\Tribunal::class, 'title' => 'Tribunaux', 'groupe' => 'Préparer', 'colonnes' => ['libelle', 'ville']],
            'communes' => ['class' => E\Commune::class, 'title' => 'Communes', 'groupe' => 'Préparer', 'colonnes' => ['nom', 'codePostal', 'canton', 'circonscription']],
            'cantons' => ['class' => E\Canton::class, 'title' => 'Cantons', 'groupe' => 'Préparer', 'colonnes' => ['nom']],
            'circonscriptions' => ['class' => E\Circonscription::class, 'title' => 'Circonscriptions', 'groupe' => 'Préparer', 'colonnes' => ['nom', 'ville', 'responsable']],
            'familles-documents' => ['class' => E\FamilleDocument::class, 'title' => 'Familles de documents', 'groupe' => 'Préparer', 'colonnes' => ['libelle']],
            'sources-donnees' => ['class' => E\SourceDonnees::class, 'title' => 'Sources de données', 'groupe' => 'Préparer', 'colonnes' => ['libelle']],
            // Entités filles, gérées depuis la fiche de leur parent
            'individus' => ['class' => E\Individu::class, 'title' => 'Individu', 'groupe' => '', 'colonnes' => ['nom', 'prenom'], 'parent' => 'dossier', 'parentSlug' => 'dossier'],
            'decisions' => ['class' => E\Decision::class, 'title' => 'Décision / recours', 'groupe' => '', 'colonnes' => ['dateCommission'], 'parent' => 'demande', 'parentSlug' => 'demande'],
            'historiques' => ['class' => E\HistoriqueIntervenant::class, 'title' => 'Intervention', 'groupe' => '', 'colonnes' => ['intervenant'], 'parent' => 'demande', 'parentSlug' => 'demande'],
            'enfants-demandes' => ['class' => E\EnfantDemande::class, 'title' => 'Enfant souhaité', 'groupe' => '', 'colonnes' => ['sexe'], 'parent' => 'demande', 'parentSlug' => 'demande'],
            'placements' => ['class' => E\Placement::class, 'title' => 'Placement', 'groupe' => '', 'colonnes' => ['nomEnfant'], 'parent' => 'demande', 'parentSlug' => 'demande'],
            'surveillances' => ['class' => E\Surveillance::class, 'title' => 'Surveillance', 'groupe' => '', 'colonnes' => ['datePrevue'], 'parent' => 'placement', 'parentSlug' => 'placement'],
        ];
    }

    /** @return array<string, array<string, array>> menus indexés par groupe */
    public static function menus(): array
    {
        $menus = [];
        foreach (self::all() as $slug => $c) {
            if ('' !== $c['groupe']) {
                $menus[$c['groupe']][$slug] = $c;
            }
        }

        return $menus;
    }
}

<?php

namespace App\Import;

use App\Entity as E;

/**
 * Correspondance entre les tables de l'ancienne base Access (copiée dans MySQL par tools/mdb2mysql.sh)
 * et les entités de la nouvelle application.
 *
 * Établie à partir des requêtes SQL, des formulaires et des états de l'ancienne application VB6 : les colonnes
 * absentes de la base source sont signalées par la commande (option --check) et simplement ignorées, sauf avec --strict.
 * Les étapes sont ordonnées : une étape ne peut référencer que des entités importées par une étape précédente.
 */
final class LegacyMapping
{
    /** @return list<Step> */
    public static function steps(): array
    {
        $ref = static fn (string $name, string $table, string $entity, string $key, string $label, array $extra = []) => new Step(
            $name, $table, $entity, [$key], ['libelle' => $label] + $extra,
        );

        return [
            // --- Référentiels
            $ref('Régimes assurance maladie', 'ASMA', E\RegimeAssuranceMaladie::class, 'ASMA_CODE', 'ASMA_DESIG'),
            $ref('Situations familiales', 'SITF', E\SituationFamiliale::class, 'SITF_CODE', 'SITF_DESIG'),
            $ref('Professions', 'PROF', E\Profession::class, 'PROF_CODE', 'PROF_DESI1'),
            $ref('Natures de demandes', 'NATU', E\NatureDemande::class, 'NATU_CODE', 'NATU_LIBELLE'),
            $ref('Types de demandes', 'TYPE', E\TypeDemande::class, 'TYPE_CODE', 'TYPE_LIBELLE'),
            $ref('Origines ethniques', 'ORET', E\OrigineEthnique::class, 'ORET_CODE', 'ORET_LIBELLE'),
            $ref('Institutions', 'INST', E\Institution::class, 'INST_CODE', 'INST_LIB'),
            $ref('Qualités', 'QUAL', E\Qualite::class, 'QUAL_CODE', 'QUAL_LIB'),
            $ref('Familles de documents', 'WFAM', E\FamilleDocument::class, 'WFAM_ID', 'WFAM_NOM'),
            $ref('Sources de données', 'WSOU', E\SourceDonnees::class, 'WSOU_ID', 'WSOU_NOM'),
            new Step('Types de recours', 'TYRE', E\TypeRecours::class, ['TYRE_NUM'], ['libelle' => 'TYRE_NOM'],
                after: static function (E\TypeRecours $t): void {
                    $t->setConseilEtat((bool) preg_match('/conseil\s+d.?[ée]tat/iu', $t->getLibelle()));
                }),
            new Step('Permanences', 'PERM', E\Permanence::class, ['PERM_CODE'], ['lieu' => 'PERM_LIEU', 'ville' => 'PERM_VILLE']),
            new Step('Paramètres de l\'échéancier', 'PARA', E\Parametre::class, ['PARA_CODE'],
                ['libelle' => 'PARA_LIB', 'nombreMois' => 'PARA_NBMOIS'],
                map: static fn (array $r) => ['id' => (int) $r['para_code']]),
            new Step('Cantons', 'CANT', E\Canton::class, ['CANT_NUM'], ['nom' => 'CANT_NOM']),
            new Step('Circonscriptions', 'CIRC', E\Circonscription::class, ['CIRC_NUM'],
                ['nom' => 'CIRC_NOM', 'adresse' => 'CIRC_ADR', 'codePostal' => 'CIRC_CP', 'ville' => 'CIRC_VILLE', 'responsable' => 'CIRC_RESP']),
            new Step('Communes', 'COMM', E\Commune::class, ['COMM_INSEE'],
                ['nom' => 'COMM_NOM', 'codePostal' => 'COMM_CP', 'codeInsee' => 'COMM_INSEE'],
                refs: ['canton' => ['CANT_NUM', E\Canton::class], 'circonscription' => ['CIRC_NUM', E\Circonscription::class]]),
            new Step('Œuvres de placement', 'OEPL', E\Oeuvre::class, ['OEPL_CODE'],
                ['designation' => 'OEPL_DESIGN', 'responsable' => 'OEPL_NOMRES', 'adresse' => 'OEPL_ADR', 'codePostal' => 'OEPL_CP',
                    'ville' => 'OEPL_VILLE', 'telephone' => 'OEPL_TEL', 'email' => 'OEPL_EMAIL', 'remarques' => 'OEPL_REM']),
            new Step('Tribunaux', 'JUGE', E\Tribunal::class, ['JUGE_NUM'],
                ['libelle' => 'JUGE_NOM', 'adresse1' => 'JUGE_ADR1', 'adresse2' => 'JUGE_ADR2', 'codePostal' => 'JUGE_CPTGI', 'ville' => 'JUGE_VILLE']),
            new Step('Intervenants', 'TRSO', E\Intervenant::class, ['TRSO_NUM'],
                ['nom' => 'TRSO_NOM', 'titre' => 'TRSO_TITR', 'telephone' => 'TRSO_TELPROF', 'actif' => 'TRSO_ACTIF'],
                refs: ['qualite' => ['QUAL_CODE', E\Qualite::class], 'institution' => ['INST_CODE', E\Institution::class], 'permanence' => ['PERM_CODE', E\Permanence::class]]),

            // --- Dossiers et individus
            new Step('Dossiers', 'DOSS', E\Dossier::class, ['DOSS_NUM'],
                ['numero' => 'DOSS_NUM', 'adresse1' => 'DOSS_ADR1', 'adresse2' => 'DOSS_ADR2', 'codePostal' => 'DOSS_CP', 'ville' => 'DOSS_VILLE',
                    'telephone' => 'DOSS_NUMTEL', 'reinstruction' => 'DOSS_REINS', 'departement' => 'DOSS_DEPT',
                    'regimeMatrimonial' => 'SITU_REGMAT', 'dateMariage' => 'SITU_DMARIA', 'lieuMariage' => 'SITU_LMARIA'],
                refs: ['commune' => ['COMM_INSEE', E\Commune::class], 'canton' => ['CANT_NUM', E\Canton::class],
                    'circonscription' => ['CIRC_NUM', E\Circonscription::class], 'situationFamiliale' => ['SITF_CODE', E\SituationFamiliale::class]]),
            new Step('Individus', 'INDI', E\Individu::class, ['INDI_NUM'],
                ['type' => 'INDI_CDFR', 'nom' => 'INDI_NOM', 'prenom' => 'INDI_PREN', 'dateNaissance' => 'INDI_DNAIS', 'lieuNaissance' => 'INDI_LNAIS',
                    'nationalite' => 'INDI_NATIO', 'dateDeces' => 'INDI_DDECE', 'dateDivorce' => 'INDI_DDIVO', 'dateVeuvage' => 'INDI_DVEUV',
                    'email' => 'INDI_EMAIL', 'adresse1' => 'INDI_ADRPRE1', 'adresse2' => 'INDI_ADRPRE2', 'codePostal' => 'INDI_CODPRE', 'ville' => 'INDI_VILPRE'],
                refs: ['dossier' => ['DOSS_NUM', E\Dossier::class], 'profession' => ['PROF_CODE', E\Profession::class],
                    'regimeAssuranceMaladie' => ['ASMA_CODE', E\RegimeAssuranceMaladie::class]],
                requires: ['DOSS_NUM']),

            // --- Demandes
            new Step('Demandes', 'DEMA', E\Demande::class, ['DEMA_NUM'],
                ['numero' => 'DEMA_NUM', 'dateDemande' => 'DEMA_DATE', 'dateSortie' => 'DEMA_DSORT', 'motifSortie' => 'DEMA_MSORT',
                    'dateSuspension' => 'DEMA_DSUSP', 'motifSuspension' => 'DEMA_MSUSP', 'assistanteMaternelle' => 'DEMA_CDDAM', 'fratrie' => 'DEMA_CFRAT',
                    'remarques' => 'DEMA_REMA', 'aActualiser' => 'DEMA_InclureReactualisation', 'dateAgrement' => 'DEMA_dtmAgrementDelivrance',
                    'provenanceAgrement' => 'DEMA_intAgrementProvenance', 'dateReceptionDossier' => 'DEMA_DREDOS', 'datePrevueDecision' => 'DEMA_DPRDEC'],
                refs: ['dossier' => ['DOSS_NUM', E\Dossier::class], 'type' => ['TYPE_CODE', E\TypeDemande::class], 'nature' => ['NATU_CODE', E\NatureDemande::class]],
                requires: ['DOSS_NUM']),

            // --- Décisions : commission initiale, recours administratifs (READ), recours contentieux (RECO)
            new Step('Décisions des commissions', 'u02_HistoCommissions', E\Decision::class, [],
                ['commissionInitiale' => 'u02_dtmHistoCommissionReunion', 'dateCommission' => 'u02_dtmHistoCommissionReunion',
                    'dateSignature' => 'u02_dtmHistoCommissionSignature', 'dateDecisionPrevue' => 'u02_dtmHistoCommissionPrevu',
                    'deliberation' => ['u02_intHistoCommissionDecision', 'deliberation'], 'remarques' => 'u02_txtHistoCommissionMotif'],
                refs: ['demande' => ['DEMA_NUM', E\Demande::class]], defaults: ['provenance' => -2], requires: ['DEMA_NUM'], mapKey: 'Decision:u02'),
            new Step('Recours administratifs', 'READ', E\Decision::class, [],
                ['commissionInitiale' => 'u02_dtmHistoCommissionReunion', 'dateCommission' => 'READ_DCOM', 'dateSignature' => 'READ_DDECI',
                    'dateDepot' => 'READ_DDREC', 'dateDecisionPrevue' => 'READ_DPREDEC', 'deliberation' => ['READ_DECIF', 'deliberation'], 'remarques' => 'READ_REFUS'],
                refs: ['demande' => ['DEMA_NUM', E\Demande::class]], defaults: ['provenance' => -1], requires: ['DEMA_NUM'], mapKey: 'Decision:read'),
            new Step('Recours contentieux', 'RECO', E\Decision::class, [],
                ['commissionInitiale' => 'u02_dtmHistoCommissionReunion', 'dateSignature' => 'RECO_dtmAgrementDelivrance', 'dateDepot' => 'RECO_DDREC',
                    'dateDecisionPrevue' => 'RECO_DPREDEC', 'deliberation' => ['RECO_DECI', 'deliberation'], 'remarques' => 'RECO_REFUS'],
                refs: ['demande' => ['DEMA_NUM', E\Demande::class], 'typeRecours' => ['TYRE_NUM', E\TypeRecours::class]],
                requires: ['DEMA_NUM'], mapKey: 'Decision:reco',
                map: static fn (array $r) => ['provenance' => (int) ($r['tyre_num'] ?? 1) ?: 1]),

            // --- Interventions des psychologues et travailleurs sociaux
            new Step('Interventions', 'u01_HistoIntervenants', E\HistoriqueIntervenant::class, [],
                ['cadre' => ['u01_intHistoIntervenantCadre', 'int'], 'dateDemande' => 'u01_dtmHistoIntervenantDemande',
                    'datePrevue' => 'u01_dtmHistoIntervenantPrevu', 'dateRecu' => 'u01_dtmHistoIntervenantRecu'],
                refs: ['demande' => ['DEMA_NUM', E\Demande::class], 'intervenant' => ['TRSO_NUM', E\Intervenant::class]],
                requires: ['DEMA_NUM', 'TRSO_NUM'], mapKey: 'Histo:u01',
                after: static function (E\HistoriqueIntervenant $h, array $r, ImportContext $ctx): void {
                    $ctx->marquerType($h->getIntervenant(), self::vrai($r['u01_boohistointervenantpsychologue'] ?? null));
                }),
            new Step('Interventions (recours, travailleur social)', 'READ', E\HistoriqueIntervenant::class, [],
                ['dateDemande' => 'READ_DDEMAS', 'datePrevue' => 'READ_DPRRECAS', 'dateRecu' => 'READ_RECAS'],
                refs: ['demande' => ['DEMA_NUM', E\Demande::class], 'intervenant' => ['TRSO_NUMAS', E\Intervenant::class]],
                defaults: ['cadre' => 4], requires: ['DEMA_NUM', 'TRSO_NUMAS'], mapKey: 'Histo:read-as',
                after: static fn (E\HistoriqueIntervenant $h, array $r, ImportContext $ctx) => $ctx->marquerType($h->getIntervenant(), false)),
            new Step('Interventions (recours, psychologue)', 'READ', E\HistoriqueIntervenant::class, [],
                ['dateDemande' => 'READ_DDEMPSY', 'datePrevue' => 'READ_DPRRECPSY', 'dateRecu' => 'READ_RECPSY'],
                refs: ['demande' => ['DEMA_NUM', E\Demande::class], 'intervenant' => ['TRSO_NUMPSY', E\Intervenant::class]],
                defaults: ['cadre' => 4], requires: ['DEMA_NUM', 'TRSO_NUMPSY'], mapKey: 'Histo:read-psy',
                after: static fn (E\HistoriqueIntervenant $h, array $r, ImportContext $ctx) => $ctx->marquerType($h->getIntervenant(), true)),

            // --- Enfants souhaités (FADE), à partir de la table ENDE ; leurs origines ethniques sont dans ORSO
            new Step('Enfants souhaités', 'FADE', E\EnfantDemande::class, [], [],
                refs: ['demande' => ['DEMA_NUM', E\Demande::class]], requires: ['DEMA_NUM', 'ENDE_NUM'],
                map: static function (array $r, ImportContext $ctx): ?array {
                    $enfant = $ctx->ligne('ENDE', 'ende_num', $r['ende_num']);
                    if (null === $enfant) {
                        return null;
                    }
                    $tranche = (string) ($enfant['ende_trage'] ?? '');
                    preg_match_all('/\d+/', $tranche, $m);
                    $origine = $ctx->ligne('ORSO', 'ende_num', $r['ende_num']);

                    return [
                        'sexe' => Transformer::apply('sexe', $enfant['ende_sexe'] ?? null),
                        'handicap' => Transformer::apply('bool', $enfant['ende_hand'] ?? null),
                        'ageMax' => $m[0] ? (int) end($m[0]) : null,
                        'origineEthnique' => $origine ? $ctx->resolve(E\OrigineEthnique::class, (string) ($origine['oret_code'] ?? '')) : null,
                    ];
                }),

            // --- Placements et surveillances
            new Step('Placements', 'PLAC', E\Placement::class, ['PLAC_NUM'],
                ['datePlacement' => 'PLAC_DATE', 'pupille' => 'PLAC_PUPILL', 'dateJugement' => 'JUAC_DJUGTPA'],
                refs: ['demande' => ['DEMA_NUM', E\Demande::class], 'oeuvre' => ['OEPL_CODE', E\Oeuvre::class], 'tribunal' => ['JUGE_NUM', E\Tribunal::class]],
                requires: ['DEMA_NUM'],
                map: static function (array $r, ImportContext $ctx): array {
                    $enfant = $ctx->ligne('ENPL', 'plac_num', $r['plac_num'] ?? null) ?? [];
                    $remarques = null;
                    if (!empty($r['plac_dcons']) && !str_starts_with((string) $r['plac_dcons'], '0000')) {
                        $remarques = 'Date de consentement : '.substr((string) $r['plac_dcons'], 0, 10);
                    }

                    return [
                        'nomEnfant' => (string) ($enfant['enpl_nom'] ?? ''),
                        'prenomEnfant' => (string) ($enfant['enpl_pren'] ?? ''),
                        'dateNaissance' => Transformer::apply('date', $enfant['enpl_dnais'] ?? null),
                        'sexe' => Transformer::apply('sexe', $enfant['enpl_sexe'] ?? null),
                        'remarques' => $remarques,
                    ];
                }),
            new Step('Surveillances', 'SUPL', E\Surveillance::class, [],
                ['datePrevue' => 'SUPL_dtmAssistanteSocialePrevu', 'dateRealisee' => 'SUPL_DDSUR', 'dateEnvoiTgi' => 'SUPL_DENVTGI'],
                refs: ['placement' => ['PLAC_NUM', E\Placement::class], 'intervenant' => ['TRSO_NUMAS', E\Intervenant::class]],
                requires: ['PLAC_NUM']),
        ];
    }

    private static function vrai(mixed $v): bool
    {
        return Transformer::apply('bool', $v) === true;
    }
}

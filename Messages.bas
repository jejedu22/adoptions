Attribute VB_Name = "modMessages"
Option Explicit

Private Const mstrNOM_MODULE As String = "modMessages"

Public Enum MessagesErreurs                         'Erreurs
    mesREGISTRE_LECTURE_ÉCRITURE = 50001
    mesCONFLIT_ACCÈS
    mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION
    mesANNULATION_SUR_DOUBLON
    mesANNULATION_SUR_CONNEXE_MANQUANT
    mesANNULATION_SUR_CONNEXE_EXISTANT
    mesENREGISTREMENT_MANQUANT
    mesRECORDSET_VIDE
    mesMESSAGE_AVEC_ICÔNE
    mesSÉLECTION_MANQUANTE
    mesSAISIE_MANQUANTE
    mesDATE_ANTÉRIEURE
    mesDATE_POSTÉRIEURE
    mesDEMANDEUR_NON_DÉCLARÉ
    mesFRATRIE_REFUSÉE
    mesENREGISTREMENT_SUPPRIMÉ
    mesDATE_SIGNATURE_ABSENTE
    mesOBLIGE
    mesDATE_RÉCEPTION_ABSENTE
    mesEXISTE_DÉJÀ
    mesIMAGE_MANQUANTE
    mesPARAMÉTRAGE_CONTENTIEUX_INCOMPLET
    mesMESSAGE_001
    mesMESSAGE_002
    mesMESSAGE_003
    mesMESSAGE_004
    mesMESSAGE_005
    mesMESSAGE_006
    mesMESSAGE_007
    mesMESSAGE_008
    mesMESSAGE_009
    mesMESSAGE_010
    mesMESSAGE_011
    mesMESSAGE_012
    mesMESSAGE_013
    mesMESSAGE_014
    mesMESSAGE_015
    mesMESSAGE_016
    mesMESSAGE_017
    mesMESSAGE_018
    mesMESSAGE_019
    mesMESSAGE_020
    mesMESSAGE_021
    mesMESSAGE_022
#If Not TablesEnLocalOuAttachées Then
    mesMESSAGE_023
#End If
End Enum
Public Enum MessagesActions                         'Actions à entreprendre suite à une erreur
    mesINTERROMPRE = 0
    mesPOURSUIVRE
End Enum
Public Enum MessagesQuestions                       'Questions
    mesSUPPRIMER = 1
    mesSUPPRIMER_INSCRIPTION_COMMISSION
    mesQUESTION_001
    mesQUESTION_002
    mesQUESTION_003
    mesQUESTION_004
    mesQUESTION_005
    mesQUESTION_006
    mesQUESTION_007
    mesQUESTION_008
    mesQUESTION_009
    mesQUESTION_010
    mesQUESTION_011
End Enum


'Déterminer l'action à entreprendre en présence d'une erreur
'   perrSource:         objet erreur
'   pstrNomProcédure:   procédure à l'origine de l'erreur
'   pstrNomModule:      module de la procédure
'   pvarArgument():     tableau d'arguments

Public Function lngAction(perrSource As ErrObject, pstrNomProcédure As String, pstrNomModule As String, pctlFocus As Control, ParamArray pvarArgument() As Variant) As MessagesActions
Dim plngCodeBouton As VbMsgBoxStyle, plngCodeIcône As VbMsgBoxStyle, plngCodeBoutonParDéfaut As VbMsgBoxStyle
Dim pstrChaînes() As String, pstrMessage As String
    
    lngAction = mesPOURSUIVRE
    
    Select Case perrSource.Number
    
'       Une erreur de lecture ou d'écriture s'est produite sur le registre
        Case mesREGISTRE_LECTURE_ÉCRITURE:
            pstrMessage = "Exécution interrompue: Une erreur " & pvarArgument(1) & " s'est produite sur le registre." & vbCrLf & vbCrLf _
                        & "Veuillez prendre note des informations ci-dessous, et contacter le service ESAS." & vbCrLf & vbCrLf _
                        & "1-  Module:        " & pstrNomModule & vbCrLf _
                        & "2-  Procédure:   " & pstrNomProcédure & vbCrLf & vbCrLf _
                        & "3-  Erreur:         " & pvarArgument(0) & " - '" & pvarArgument(2) & "'"
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbCritical
            lngAction = mesINTERROMPRE

'       Une opération sur la base a engendré un conflit d'accès inattendu
        Case mesCONFLIT_ACCÈS:
            pstrMessage = "Exécution interrompue: Une opération sur la base a engendré un conflit d'accès." & vbCrLf & vbCrLf _
                                            & "Vous pouvez tenter de répéter l'opération après avoir relancé l'application." & vbCrLf & vbCrLf _
                                            & "Veuillez cependant prendre note des informations ci-dessous et contacter le service ESAS." & vbCrLf & vbCrLf _
                                            & "1-  Module:        " & pstrNomModule & vbCrLf _
                                            & "2-  Procédure:   " & pstrNomProcédure & vbCrLf & vbCrLf _
                                            & "3-  Conflit:          " & pvarArgument(0)
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbCritical
            lngAction = mesINTERROMPRE
                            
'       L'utilisateur demande l'annulation du traitement en raison d'un verrouillage ou de la modification des données par un autre utilisateur
        Case mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION:
            pstrChaînes = Split(pvarArgument(0), ",")
            If UBound(pstrChaînes) = 0 Then pstrMessage = pstrChaînes(0) Else pstrMessage = pstrChaînes(0) & " est annulé" & pstrChaînes(1) & "."
            plngCodeBouton = vbOKOnly: plngCodeIcône = pvarArgument(1)

'       L'ajout (ou la mise à jour) a été refusé(e) pour éviter la création d'un doublon
        Case mesANNULATION_SUR_DOUBLON:
            pstrMessage = "Exécution interrompue: " & pvarArgument(0) & " d'un enregistrement a été refusé" & pvarArgument(1) & " pour éviter la création d'un doublon." & vbCrLf & vbCrLf _
                        & "Vous pouvez tenter de répéter l'opération après avoir relancé l'application." & vbCrLf & vbCrLf _
                        & "Si le problème persiste, veuillez prendre note des informations ci-dessous et contacter le service ESAS." & vbCrLf & vbCrLf _
                        & "1-  Module:        " & pstrNomModule & vbCrLf _
                        & "2-  Procédure:   " & pstrNomProcédure & vbCrLf & vbCrLf _
                        & "3-  Requête:     " & pvarArgument(2)
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbCritical
            lngAction = mesINTERROMPRE

'       L'ajout (ou la mise à jour) a été refusé(e) du fait d'un enregistrement connexe manquant
        Case mesANNULATION_SUR_CONNEXE_MANQUANT:
            pstrMessage = "Exécution interrompue: " & pvarArgument(0) & " d'un enregistrement a été refusé" & pvarArgument(1) & " du fait d'un enregistrement connexe manquant." & vbCrLf & vbCrLf _
                        & "Vous pouvez tenter de répéter l'opération après avoir relancé l'application." & vbCrLf & vbCrLf _
                        & "Si le problème persiste, veuillez prendre note des informations ci-dessous et contacter le service ESAS." & vbCrLf & vbCrLf _
                        & "1-  Module:            " & pstrNomModule & vbCrLf _
                        & "2-  Procédure:       " & pstrNomProcédure & vbCrLf & vbCrLf _
                        & "3-  Table:              " & pvarArgument(2) & vbCrLf & vbCrLf _
                        & "4-  Requête:         " & pvarArgument(3)
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbCritical
            lngAction = mesINTERROMPRE

'       La suppression (ou la mise à jour) a été refusée du fait d'un enregistrement connexe existant
        Case mesANNULATION_SUR_CONNEXE_EXISTANT:
            pstrMessage = "Exécution interrompue: La " & pvarArgument(0) & " d'un enregistrement a été refusée du fait d'un enregistrement connexe existant." & vbCrLf & vbCrLf _
                        & "Veuillez prendre note des informations ci-dessous, et contacter le service ESAS." & vbCrLf & vbCrLf _
                        & "1-  Module:            " & pstrNomModule & vbCrLf _
                        & "2-  Procédure:       " & pstrNomProcédure & vbCrLf & vbCrLf _
                        & "3-  Table:              " & pvarArgument(1) & vbCrLf & vbCrLf _
                        & "4-  Requête:         " & pvarArgument(2)
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbCritical
            lngAction = mesINTERROMPRE

'       Un enregistrement manque dans la base
        Case mesENREGISTREMENT_MANQUANT:
            pstrMessage = "Exécution interrompue: Un enregistrement manque dans la base." & vbCrLf & vbCrLf _
                        & "Veuillez prendre note des informations ci-dessous, et contacter le service ESAS." & vbCrLf & vbCrLf _
                        & "1-  Module:        " & pstrNomModule & vbCrLf _
                        & "2-  Procédure:   " & pstrNomProcédure & vbCrLf & vbCrLf _
                        & "3-  Table:           " & pvarArgument(0) & vbCrLf _
                        & "4-  Clé:               " & pvarArgument(1)
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbCritical
            lngAction = mesINTERROMPRE

'       Le recordset est vide
        Case mesRECORDSET_VIDE:
            pstrMessage = "Exécution interrompue: " & pvarArgument(0) & vbCrLf & vbCrLf _
                        & "Vous pouvez tenter de répéter l'opération après avoir relancé l'application." & vbCrLf & vbCrLf _
                        & "Si le problème persiste, veuillez prendre note des informations ci-dessous et contacter le service ESAS." & vbCrLf & vbCrLf _
                        & "1-  Module:        " & pstrNomModule & vbCrLf _
                        & "2-  Procédure:   " & pstrNomProcédure & vbCrLf & vbCrLf _
                        & "3-  Requête:     " & pvarArgument(1)
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbCritical
            lngAction = mesINTERROMPRE

'       Afficher un message avec son icône
        Case mesMESSAGE_AVEC_ICÔNE:
            pstrMessage = pvarArgument(0)
            plngCodeBouton = vbOKOnly: plngCodeIcône = pvarArgument(1)

'       Une sélection n'est pas effectuée
        Case mesSÉLECTION_MANQUANTE:
            If Not Right(pvarArgument(0), Len(vbCrLf)) = vbCrLf Then pvarArgument(0) = pvarArgument(0) & " "
            pstrMessage = ": La sélection " & pvarArgument(0) & "est obligatoire."
            If UBound(pvarArgument) = 1 Then pstrMessage = pvarArgument(1) & pstrMessage Else pstrMessage = "Validation refusée" & pstrMessage
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       Une saisie n'est pas effectuée
        Case mesSAISIE_MANQUANTE:
            If Not Right(pvarArgument(0), Len(vbCrLf)) = vbCrLf Then pvarArgument(0) = pvarArgument(0) & " "
            pstrMessage = ": La saisie " & pvarArgument(0) & "est obligatoire."
            If UBound(pvarArgument) = 1 Then pstrMessage = pvarArgument(1) & pstrMessage Else pstrMessage = "Validation refusée" & pstrMessage
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       Une date est antérieure à une autre
        Case mesDATE_ANTÉRIEURE:
            If Not Right(pvarArgument(0), Len(vbCrLf)) = vbCrLf Then pvarArgument(0) = pvarArgument(0) & " "
            If Not Left(pvarArgument(1), Len(vbCrLf)) = vbCrLf Then pvarArgument(1) = " " & pvarArgument(1)
            pstrMessage = ": La date " & pvarArgument(0) & "est antérieure à la date" & pvarArgument(1) & "."
            If UBound(pvarArgument) = 2 Then pstrMessage = pvarArgument(2) & pstrMessage Else pstrMessage = "Validation refusée" & pstrMessage
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       Une date est postérieure à une autre
        Case mesDATE_POSTÉRIEURE:
            If Not Right(pvarArgument(0), Len(vbCrLf)) = vbCrLf Then pvarArgument(0) = pvarArgument(0) & " "
            If Not Left(pvarArgument(1), Len(vbCrLf)) = vbCrLf Then pvarArgument(1) = " " & pvarArgument(1)
            pstrMessage = ": La date " & pvarArgument(0) & "est postérieure à la date" & pvarArgument(1) & "."
            If UBound(pvarArgument) = 2 Or UBound(pvarArgument) = 3 Then pstrMessage = pvarArgument(2) & pstrMessage Else pstrMessage = "Validation refusée" & pstrMessage
            If UBound(pvarArgument) = 3 Then pstrMessage = Replace(pstrMessage, "postérieure", pvarArgument(3))
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       Le demandeur n'est pas déclaré
        Case mesDEMANDEUR_NON_DÉCLARÉ:
            pstrMessage = "Le demandeur n'est pas déclaré."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbInformation

'       La fratrie n'est pas autorisée
        Case mesFRATRIE_REFUSÉE:
            pstrMessage = pvarArgument(0) & " refusée: La saisie d'une fratrie n'est pas autorisée."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       Un enregistrement a été supprimé par un autre utilisateur
        Case mesENREGISTREMENT_SUPPRIMÉ:
            If Not Right(pvarArgument(1), Len(vbCrLf)) = vbCrLf Then pvarArgument(1) = pvarArgument(1) & " "
            pstrMessage = pvarArgument(0) & " impossible: " & pvarArgument(1) & "n'existe plus."
            plngCodeBouton = vbOKOnly: If UBound(pvarArgument) = 2 Then plngCodeIcône = pvarArgument(2) Else plngCodeIcône = vbExclamation

'       La date de signature de la décision prise en commission n'est pas saisie
        Case mesDATE_SIGNATURE_ABSENTE:
            pstrMessage = "La date de signature du " & pvarArgument(0) & " décidé par la commission" & vbCrLf _
                        & "n'étant pas saisie, toute validation est impossible."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbInformation

'       La saisie d'une information oblige la saisie d'une autre
        Case mesOBLIGE:
            If Not Right(pvarArgument(0), Len(vbCrLf)) = vbCrLf Then pvarArgument(0) = pvarArgument(0) & " "
            If Not Left(pvarArgument(1), Len(vbCrLf)) = vbCrLf Then pvarArgument(1) = " " & pvarArgument(1)
            pstrMessage = ": " & pvarArgument(0) & "oblige" & pvarArgument(1) & "."
            If UBound(pvarArgument) = 2 Then pstrMessage = pvarArgument(2) & pstrMessage Else pstrMessage = "Validation refusée" & pstrMessage
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       La date de réception du dossier n'est pas saisie
        Case mesDATE_RÉCEPTION_ABSENTE:
            pstrMessage = "Validation refusée: L'absence d'une date de réception pour le dossier interdit " & pvarArgument(0) & "."
            plngCodeBouton = vbOKOnly: plngCodeIcône = IIf(pvarArgument(0) = "interdite", vbCritical, vbExclamation)

'       L'entité existante déjà
        Case mesEXISTE_DÉJÀ:
            pstrMessage = pvarArgument(0) & " refusée: " & pvarArgument(1) & " existe déjà"
            If UBound(pvarArgument) = 2 Then pstrMessage = pstrMessage & " " & pvarArgument(2)
            pstrMessage = pstrMessage & "."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       Menu: Un fichier image est manquant
        Case mesIMAGE_MANQUANTE:
            pstrMessage = "Exécution interrompue: Un fichier d'images n'a pu être chargé." & vbCrLf & vbCrLf _
                        & "Vous pouvez tenter de répéter l'opération après avoir relancé l'application." & vbCrLf & vbCrLf _
                        & "Si le problème persiste, veuillez prendre note des informations ci-dessous et contacter le service ESAS." & vbCrLf & vbCrLf _
                        & "1-  Module:        " & pstrNomModule & vbCrLf _
                        & "2-  Procédure:   " & pstrNomProcédure & vbCrLf & vbCrLf _
                        & "3-  Fichier:         " & pvarArgument(0) & vbCrLf _
                        & "4-  Erreur:         " & pvarArgument(1)
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbCritical
            lngAction = mesINTERROMPRE

'       Frm_Recours: Cour Administrative d'Appel non paramétrée
        Case mesPARAMÉTRAGE_CONTENTIEUX_INCOMPLET:
            pstrMessage = "Fonctionnalité refusée: Le paramétrage des différents types de recours contentieux n'est pas complet." & vbCrLf & vbCrLf _
                        & "La solution consiste à vérifier le paramétrage et l'orthographe exacte des libellés (y compris leur casse). Les libellés suivants doivent apparaître:" & vbCrLf & vbCrLf _
                        & "1-  " & pvarArgument(0) & vbCrLf _
                        & "2-  " & pvarArgument(1) & vbCrLf _
                        & "3-  " & pvarArgument(2)
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       modDAO: La clé 'Adoptions' n'existe pas dans le registre
        Case mesMESSAGE_001:
            pstrMessage = "Exécution interrompue: Le driver ODBC n'existe pas pour la base Adoptions." & vbCrLf & vbCrLf _
                        & "Veuillez contacter le service ESAS."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbCritical
            lngAction = mesINTERROMPRE

'       frm_RecalculEcheancier: Nature des dates recalculées
        Case mesMESSAGE_002:
            pstrMessage = "Le traitement a procédé au recalcul des dates suivantes:" & vbCrLf _
                        & pvarArgument(0) & "."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbInformation

'       frmDemandesÀRéactualiser: Aucune demande à réactualiser
        Case mesMESSAGE_003:
            pstrMessage = "Aucune demande n'est à réactualiser."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbInformation

'       frm_PlacementGestion: Suppression du placement
        Case mesMESSAGE_004:
            pstrMessage = "Le placement a été supprimé."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbInformation

'       frm_DecisionCommission: Date de référence non définie
        Case mesMESSAGE_005:
            pstrMessage = "La date de réception du dossier n'étant pas saisie, toute validation est impossible." & vbCrLf & vbCrLf _
                        & "Remarque: Si la date de réception est bien saisie, c'est que la date de référence servant au calcul de l'échéancier" & vbCrLf _
                        & "n'a pas été positionnée dans une version antérieure du logiciel Adoptions. La solution consiste alors à valider" & vbCrLf _
                        & "le recalcul de l'échéancier à partir des informations proposées."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbInformation

'       frm_DecisionCommission: Aucun profil d'enfant saisi
        Case mesMESSAGE_006:
            pstrMessage = pvarArgument(0) & " refusée: La saisie d'un profil d'enfant est obligatoire," & vbCrLf _
                        & "c'est-à-dire, au minimum, d'un enfant associé à une origine ethnique."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       frm_DecisionCommission: Décision de retrait sans agrément préalable
        Case mesMESSAGE_007:
            pstrMessage = pvarArgument(0) & " refusée: Elle conduirait à l'incohérence des décisions." & vbCrLf & vbCrLf _
                        & "La décision de retrait du " & pvarArgument(1) & " ne fait suite à aucun agrément."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       Frm_Recours: Recours contentieux épuisés
        Case mesMESSAGE_008:
            pstrMessage = "Création refusée: Tous les recours contentieux sont épuisés."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       Frm_Recours: Recours postérieur
        Case mesMESSAGE_009:
            pstrMessage = "Suppression refusée: Un autre recours a été déposé après celui-ci." & vbCrLf & vbCrLf _
                        & "La solution consiste à supprimer ce recours avant de supprimer celui-ci."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       Frm_Recours: Aucune décision prise à l'occasion du précédent recours contentieux
        Case mesMESSAGE_010:
            pstrMessage = "Création refusée: Aucune décision n'a été prise à l'occasion du " & pvarArgument(0) & " recours."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       frm_DemandeAdoptionInstruction: Mise à jour des informations relatives à un recours administratif ou à une surveillance de placement interdite sur l'instruction
        Case mesMESSAGE_011:
            pstrMessage = "Fonctionnalité refusée: Les informations relatives à " & pvarArgument(0) & vbCrLf _
                        & "ne peuvent être " & pvarArgument(1) & " qu'à l'aide de la fenêtre réservée à leur traitement."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       modAdoptions: L'intervenant est inconnu
        Case mesMESSAGE_012:
            pstrMessage = "L'intervenant est inconnu." & vbCrLf _
                        & "[Numéro=" & pvarArgument(0) & "]"
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbInformation

'       frm_DossierGestion: La demande n'est pas créée
        Case mesMESSAGE_013:
            pstrMessage = "Fonctionnalité refusée: Aucune demande n'a été créée."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       modAdoptions: Date de dépôt de recours incohérente
        Case mesMESSAGE_014:
            pstrMessage = pvarArgument(0) & " refusée: La date de dépôt n'est pas cohérente avec celle des autres recours."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       modAdoptions: Recours administratif ou contentieux sans décision préalable de refus ou de retrait
        Case mesMESSAGE_015:
            pstrMessage = pvarArgument(0) & " refusée: Elle conduirait à l'incohérence des décisions." & vbCrLf & vbCrLf _
                        & "Le recours " & pvarArgument(1) & pvarArgument(2) & " ne fait suite à aucune décision de refus ou de retrait."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       modAdoptions: Création d'une nouvelle commission en l'absence d'agrément
        Case mesMESSAGE_016:
            pstrMessage = pvarArgument(0) & " refusée: Elle conduirait à l'incohérence des décisions." & vbCrLf & vbCrLf _
                        & pvarArgument(1) & " interdit la création d'une nouvelle commission le " & pvarArgument(2) & "."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       modAdoptions: Succession de recours contentieux incohérente
        Case mesMESSAGE_017:
            pstrMessage = pvarArgument(0) & " refusée: Elle conduirait à l'incohérence des décisions." & vbCrLf & vbCrLf _
                        & "La succession des recours contentieux ne respecte pas l'ordre imposé par le paramétrage."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       modAdoptions: Recours contentieux sans décision du recours précédent
        Case mesMESSAGE_018:
            pstrMessage = pvarArgument(0) & " refusée: Elle conduirait à l'incohérence des décisions." & vbCrLf & vbCrLf _
                        & "L'existence d'un recours déposé le " & pvarArgument(1) & " oblige la saisie d'une décision pour celui déposé le " & pvarArgument(2) & "."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       modAdoptions: Recours sans date de signature pour la décision initiale
        Case mesMESSAGE_019:
            pstrMessage = pvarArgument(0) & " refusée: Elle conduirait à l'incohérence des décisions." & vbCrLf & vbCrLf _
                        & "L'absence d'une date de signature pour le " & pvarArgument(1) & " décidé par la commission initiale interdit le dépôt d'un recours le " & pvarArgument(2) & "."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       frm_PlacementAjoutTS: Date de réception du dossier non saisie
        Case mesMESSAGE_020:
            pstrMessage = "La date de réception du dossier n'étant pas saisie, toute validation est impossible."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbInformation

'       frm_Echéancier: Date invalide
        Case mesMESSAGE_021:
            pstrMessage = "La date n'est pas une date valide."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation

'       modDAO: Fichier inexistant ou inaccessible
        Case mesMESSAGE_022:
            pstrMessage = "Exécution interrompue: " & pvarArgument(0) & " est inexistant" & pvarArgument(1) & " ou inaccessible." & vbCrLf & vbCrLf _
                        & "Vous pouvez tenter de répéter l'opération après avoir relancé l'application." & vbCrLf & vbCrLf _
                        & "Si le problème persiste, veuillez prendre note des informations ci-dessous et contacter le service ESAS." & vbCrLf & vbCrLf _
                        & "1-  Module:        " & pstrNomModule & vbCrLf _
                        & "2-  Procédure:   " & pstrNomProcédure & vbCrLf & vbCrLf _
                        & "3-  Fichier:         " & pvarArgument(2) & vbCrLf _
                        & "4-  Erreur:         " & pvarArgument(3)
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbCritical
            lngAction = mesINTERROMPRE

#If Not TablesEnLocalOuAttachées Then
'       modDAO: Création impossible du fait de l'ouverture, par un autre utilisateur, de la base ou de l'objet en mode lecture seule (Erreur DAO 1809)
        Case mesMESSAGE_023:
            pstrMessage = "Les données sur lesquelles vous travaillez sont actuellement verrouillées" & vbCrLf _
                        & "par un autre utilisateur, ce qui interdit pour le moment toute création."
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbExclamation
#End If

'       Le code a généré une erreur inattendue
        Case Else:
            pstrMessage = "Exécution interrompue: Une erreur met fin au traitement." & vbCrLf & vbCrLf _
                        & "Veuillez prendre note des informations ci-dessous, et contacter le service ESAS." & vbCrLf & vbCrLf _
                        & "1-  Module:        " & pstrNomModule & vbCrLf _
                        & "2-  Procédure:   " & pstrNomProcédure & vbCrLf & vbCrLf _
                        & "3-  Erreur:         " & Format(perrSource.Number) & " - " & perrSource.Description
            plngCodeBouton = vbOKOnly: plngCodeIcône = vbCritical
            lngAction = mesINTERROMPRE
                            
    End Select
    
    Beep
    MsgBox pstrMessage, plngCodeBouton + plngCodeIcône + plngCodeBoutonParDéfaut, "Adoptions - " & Choose(plngCodeIcône / 16, IIf(lngAction = mesINTERROMPRE, "Message critique", "Interdiction"), "", "Avertissement", "Information")

'   S'il s'agit d'une erreur ODBC
    If Err.Number = daoERREUR_03146 Then
        pstrMessage = modCommun.strChaîneErreursODBC()
        MsgBox "Veuillez prendre note de ces autres informations :" & vbCrLf & vbCrLf & pstrMessage, vbCritical + vbOKOnly, "Message critique"
    End If
    
'   Donner le focus à un contrôle si demandé
    If Not pctlFocus Is Nothing Then
        If pctlFocus.Enabled Then pctlFocus.SetFocus
    End If

#If Débogage Then
    If lngAction = mesINTERROMPRE Then Stop
    lngAction = mesPOURSUIVRE
#End If
    
End Function


'Retourner la réponse de l'utilisateur à une question après l'avoir affichée
'   Arguments:  plngNuméroQuestion: numéro de la question
'               pvarArgument():     tableau de type ParamArray constitué de chaque chaîne implémentant le libellé de la question
Public Function lngRéponseUtilisateur(plngNuméroQuestion As MessagesQuestions, ParamArray pvarArgument() As Variant) As VbMsgBoxResult
Const pkstrNomProcédure As String = "lngRéponseUtilisateur"
Dim pstrQuestion As String, plngCodeBoutons As Long
On Error GoTo Err_lngRéponseUtilisateur

    Select Case plngNuméroQuestion
    
'       Confirmer la suppression
        Case mesSUPPRIMER:
            pstrQuestion = "Êtes-vous certain de vouloir supprimer " & pvarArgument(0) & " ?"
            plngCodeBoutons = vbYesNo + vbDefaultButton2

'       Confirmer la suppression de l'inscription du dossier à la commission
        Case mesSUPPRIMER_INSCRIPTION_COMMISSION:
            pstrQuestion = "Êtes-vous certain de vouloir désinscrire le dossier de la commission du " & pvarArgument(0) & " ?"
            plngCodeBoutons = vbYesNo + vbDefaultButton2

'       modDAO: Proposer la poursuite de l'opération sur verrouillage (Erreurs DAO 1304, 3035, 3049, 3260, 3261)
        Case mesQUESTION_001:
            pstrQuestion = "Les données sur lesquelles vous travaillez sont actuellement verrouillées"
            If UBound(pvarArgument) = 1 Then pstrQuestion = pstrQuestion & vbCrLf & "par l'utilisateur '" & pvarArgument(0) & "' sur l'ordinateur '" & pvarArgument(1) & "'." Else pstrQuestion = pstrQuestion & " par un autre utilisateur."
            pstrQuestion = pstrQuestion & vbCrLf & vbCrLf _
                         & "Voulez-vous tenter de répéter l'opération ?"
            plngCodeBoutons = vbYesNo + vbDefaultButton1
            
'       modDAO: Proposer d'écraser les modifications apportées entre-temps par un autre utilisateur
        Case mesQUESTION_002:
            pstrQuestion = "Certaines données ont été modifiées par un utilisateur depuis que vous les utilisez." & vbCrLf _
                         & "Si vous enregistrez vos modifications, vous écraserez les siennes." & vbCrLf & vbCrLf _
                         & "Êtes-vous certain de vouloir poursuivre la mise à jour ?"
            plngCodeBoutons = vbYesNo + vbDefaultButton2

'       frm_PlacementGestion: Proposer la création du placement
        Case mesQUESTION_003:
            pstrQuestion = "La création du placement est nécessaire à la poursuite du traitement." & vbCrLf & vbCrLf _
                         & "Confirmez-vous la création ?"
            plngCodeBoutons = vbYesNo + vbDefaultButton1

'       frm_Individu: Informer de l'existence de demandeurs portant mêmes nom et prénom
        Case mesQUESTION_004:
            pstrQuestion = "Le(s) dossier(s) ci-dessous est(sont) lié(s) à un demandeur" & vbCrLf _
                         & "portant les mêmes nom et prénom :" & vbCrLf _
                         & pvarArgument(0) & vbCrLf & vbCrLf _
                         & "Êtes-vous certain de vouloir valider votre saisie ?"
            plngCodeBoutons = vbYesNo + vbDefaultButton2

'       frm_DecisionCommission: Proposer l'enregistrement d'une nouvelle commission
        Case mesQUESTION_005:
            pstrQuestion = "Souhaitez-vous enregistrer une nouvelle commission, ou mettre à jour" & vbCrLf _
                         & "la commission précédemment enregistrée au " & pvarArgument(0) & " ?" & vbCrLf & vbCrLf _
                         & "- Oui   =  Créer" & vbCrLf _
                         & "- Non  =  Mettre à jour" & vbCrLf
            plngCodeBoutons = vbYesNoCancel + vbDefaultButton3

'       frm_DecisionCommission: Proposer les souhaits exprimés lors de la demande
        Case mesQUESTION_006:
            pstrQuestion = "Souhaitez-vous reprendre les souhaits exprimés lors de la demande ?"
            plngCodeBoutons = vbYesNo + vbDefaultButton1

'       frm_DemandeAdoptionInstruction: Proposer la mise à jour des informations relatives aux intervenants présentes sur la demande
        Case mesQUESTION_007:
            pstrQuestion = "Les informations présentes sur l'instruction ne correspondent pas à l'historique des intervenants." & vbCrLf & vbCrLf _
                         & "Souhaitez-vous leur mise à jour ?"
            plngCodeBoutons = vbYesNo + vbDefaultButton1

'       frm_DemandeAdoptionInstruction: Proposer la validation de l'instruction
        Case mesQUESTION_008:
            pstrQuestion = "Le traitement d'un intervenant suppose la présence d'une date de réception pour le dossier qui," & vbCrLf _
                         & "après validation de l'instruction, constituera la date de référence pour le calcul de l'échéancier." & vbCrLf & vbCrLf _
                         & "Souhaitez-vous valider l'instruction maintenant ?"
            plngCodeBoutons = vbYesNo + vbDefaultButton1

'       modAdoptions: Proposer la mise à jour de l'historique des intervenants
        Case mesQUESTION_009:
            pstrQuestion = "Les informations relatives aux intervenants ne correspondent pas à leur historique." & vbCrLf & vbCrLf _
                         & "Souhaitez-vous la mise à jour de l'historique ?"
            plngCodeBoutons = vbYesNo + vbDefaultButton1

'       frm_DemandeAdoptionSuspension: Confirmer la saisie d'une date de suspension antérieure à la précédente reprise
        Case mesQUESTION_010:
            pstrQuestion = "La date de suspension est antérieure ou égale à la date de la précédente reprise." & vbCrLf & vbCrLf _
                         & "Confirmez-vous la validation ?"
            plngCodeBoutons = vbYesNo + vbDefaultButton2

'       frm_Echéancier: Proposer la date du jour comme base de calcul pour l'échéancier
        Case mesQUESTION_011:
            pstrQuestion = "Souhaitez-vous prendre la date du jour comme base de calcul pour l'échancier ?"
            plngCodeBoutons = vbYesNo + vbDefaultButton1

    End Select
    
    plngCodeBoutons = plngCodeBoutons + vbQuestion
    
    Beep
    lngRéponseUtilisateur = MsgBox(pstrQuestion, plngCodeBoutons, "Question")
    
    Exit Function
    
Err_lngRéponseUtilisateur:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Function


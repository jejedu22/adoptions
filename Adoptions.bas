Attribute VB_Name = "modAdoptions"
Option Explicit

Private Const mstrNOM_MODULE As String = "modAdoptions"

Public Enum AdoptionsIndexPictureClip           'Index des contrôles PictureClip
    adpTEXTE_IMAGE = 0
    adp16x16
    adp32x32
End Enum
Public Enum AdoptionsImages                     'Index des images dans les contrôles PictureClip et ImageList
    adpTEXTE_RAFRAÎCHIR = 0
    adpTEXTE_ENVOI_DOSSIER_TGI
    adpTEXTE_HISTORIQUE
    adpANNULER = 0
    adpVALIDER
    adpFERMER
    adpCRÉER = 0
    adpMODIFIER
    adpSUPPRIMER
    adpRAFRAÎCHIR
    adpENVOYER
    adpHISTORIQUE
    adpINTERVENANTS
End Enum
Public Enum AdoptionsDélibérations                  'Valeurs des délibérations
    adpINDÉTERMINÉ = 0
    adpREFUS
    adpRETRAIT
    adpACCORD
End Enum
Public Enum AdoptionsIntervenants                   'Types d'intervenants
    adpPSYCHOLOGUE = 2
    adpTRAVAILLEUR_SOCIAL = 4
End Enum
Public Enum AdoptionsCadresInterventions            'Cadres d'interventions
    adpINDÉTERMINÉ = 0
    adpDEMANDE_INITIALE
    adpSECONDE_ÉVALUATION
    adpREPRISE
    adpRECOURS_ADMINISTRATIF
    adpSURVEILLANCE_PLACEMENT
End Enum
Public Enum AdoptionsProvenances                    'Provenances
    adpDEMANDE_INITIALE = -2
    adpRECOURS_ADMINISTRATIF
    adpINDÉTERMINÉE
End Enum
Public Enum AdoptionsSexes                          'Sexes
    adpINDÉTERMINÉ = 1
    adpMASCULIN
    adpFÉMININ
End Enum
Public Enum AdoptionsFeuilles                       'Positions des feuilles
    adpMENU_HAUTEUR = 705
    adpDEMANDES_À_RÉACTUALISER_HAUTEUR = 2190
    adpDEMANDES_À_RÉACTUALISER_DÉCALAGE_BAS = 250
End Enum
Public Enum AdoptionsDivers                         'Constantes diverses
    adpAUCUNE_SÉLECTION = 50000
    adpNOMBRE_ORIGINES_ETHNIQUES = 4
End Enum

Public Const gstrDOSSIER_IMAGES As String = "Images", gstrICONE_LISTE As String = "Liste.ico"
Public Const gstrCONSULTATION_LIMITÉE As String = "La consultation est limitée: Les informations sont incomplètes.", gstrVALIDATION_ANNULÉE As String = "La validation est annulée.", gstrLECTURE_HISTORIQUE_ANNULÉE As String = "La lecture de l'historique est annulée."
Public Const gstrDEMANDE_SUPPRIMÉE As String = "La demande n'existe plus."
Public Const gstrTYPE_DEMANDEUR As String = "DE"
Public Const gstrACCORD As String = "A", gstrREFUS As String = "R"

Public gclsDélibérations As clsCodesLibellés, gclsCadresInterventions As clsCodesLibellés


'Actualiser les données relatives à l'agrément d'une demande
'   plngNuméroDemande:          numéro de la demande
'   pstrMessageSiVerrouillage:  message à afficher si les données sont verrouillées par une transaction

Public Sub ActualiserDonnéesAgrément(plngNuméroDemande As Long, pstrMessageSiVerrouillage As String)
Const pkstrNomProcédure As String = "ActualiserDonnéesAgrément"
Dim prstDécisions As Recordset
Dim pvarPremièreCommissionÀPrendreDécision As Variant, pvarAgrémentDate As Variant, pvarAgrémentProvenance As Variant, pvarDateRetraitInitiale As Variant, pvarDateRetrait As Variant
Dim plngDécision As AdoptionsDélibérations
On Error GoTo Err_ActualiserDonnéesAgrément

'   Lire l'ensemble des décisions relatives à une demande
'       Notes:
'           Voir les règles de gestion définies en entête de la feuille frm_DecisionCommission
'           Les décisions sont triées de façon ascendante par date de commission et provenance (Demande initiale, puis Recours administratif / Tribunal Administratif / Cour Administrative d'Appel / Conseil d'État)
'           Il existe 3 types d'enregistrements, chacun implémentant de façon différente les colonnes DateCommission et DateSignature:
'           - enregistrement issu de la saisie d'une commission:
'               DateCommission: date de la commission       (n'est jamais nulle)
'               DateSignature:  date de la signature        (peut être nulle)
'               DateDepot:      NULL
'           - enregistrement issu de la saisie d'un recours administratif:
'               DateCommission: date de la commission       (peut être nulle)
'               DateSignature:  date de la signature        (peut être nulle)
'                   dans le cas où la décision est prise par le Président du Conseil Général sans consultation d'une commission, la date de la commission est nulle et la date de la décision correspond à la date de la signature
'               DateDepot:      date du dépôt               (n'est jamais nulle)
'           - enregistrement issu de la saisie d'un recours contentieux:
'               DateCommission: NULL
'               DateSignature:  date de l'agrément          (peut être nulle)
'               DateDepot:      date du dépôt               (n'est jamais nulle)
    Set prstDécisions = modDAO.rstSQLExtraction(daoDÉCISIONS, True, pstrMessageSiVerrouillage, plngNuméroDemande)
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub
    
    pvarPremièreCommissionÀPrendreDécision = Null: pvarAgrémentDate = Null: pvarAgrémentProvenance = Null: pvarDateRetraitInitiale = Null: pvarDateRetrait = Null
    
    With prstDécisions
    
'       Pour chaque décision
        While Not .EOF
        
            Select Case modCommun.varValeurFormatéeSiNull(![Decision], vbString)
                Case CStr(AdoptionsDélibérations.adpINDÉTERMINÉ), "":   plngDécision = AdoptionsDélibérations.adpINDÉTERMINÉ
                Case CStr(adpREFUS), gstrREFUS:                         plngDécision = adpREFUS
                Case CStr(adpRETRAIT):                                  plngDécision = adpRETRAIT
                Case CStr(adpACCORD), gstrACCORD:                       plngDécision = adpACCORD
            End Select
            
'           Si la décision provient d'une commission alors que la première commission à prendre une décision n'a pas encore été identifiée
            If ![Provenance] = AdoptionsProvenances.adpDEMANDE_INITIALE And IsNull(pvarPremièreCommissionÀPrendreDécision) Then
'               Si la décision est définie, référencer la commission comme la première à prendre une décision
                If Not plngDécision = AdoptionsDélibérations.adpINDÉTERMINÉ Then pvarPremièreCommissionÀPrendreDécision = ![CommissionInitiale]
            End If

'           Si la décision fait référence à la première commission à prendre une décision
            If ![CommissionInitiale] = pvarPremièreCommissionÀPrendreDécision Then
            
                Select Case plngDécision
                
'                   Si elle n'est pas déterminée
                    Case AdoptionsDélibérations.adpINDÉTERMINÉ
                
'                       S'il ne s'agit pas d'une décision du Conseil d'État
                        If Not ![Provenance] = gclsRecoursContentieux.lngConseilÉtat Then
                            pvarAgrémentDate = Null: pvarAgrémentProvenance = Null
                        End If
                    
'                   S'il s'agit d'un refus
                    Case adpREFUS
                    
                        pvarAgrémentDate = Null: pvarAgrémentProvenance = Null
                    
'                   S'il s'agit d'un accord
                    Case adpACCORD
                    
'                       S'il ne s'agit pas d'une décision du Conseil d'État ou si la décision précédente n'était pas favorable
                        If Not ![Provenance] = gclsRecoursContentieux.lngConseilÉtat Or IsNull(pvarAgrémentDate) Then
'                           Faire de ...
'                               - la date de réunion de la commission (cas de la commission initiale)
'                               - la date de réunion de la commission, ou la date de signature si la décision est prise par le Président du Conseil Général sans consultation d'une commission (cas du recours administratif)
'                               - la date de délivrance de l'agrément par le Président du Conseil Général (cas du recours contentieux)
'                           ... la date de délivrance de l'agrément initial
                            pvarAgrémentDate = IIf(IsNull(![DateCommission]), ![DateSignature], ![DateCommission])
'                           Si la date de délivrance de l'agrément initial est connue, référencer sa provenance (commission initiale, recours administratif, ou l'un quelconque des recours contentieux)
                            pvarAgrémentProvenance = IIf(IsNull(pvarAgrémentDate), Null, ![Provenance])
                        End If
                    
                End Select
                
'           Sinon
            Else
            
'               S'il s'agit d'un retrait (il ne peut être décidé que par une commission initiale)
                If plngDécision = adpRETRAIT Then
'                   Le référencer
                    pvarDateRetraitInitiale = ![DateCommission]
                End If
                
'               S'il s'agit d'un retrait ou si le recours se rapporte à un retrait
                If Not IsNull(pvarDateRetraitInitiale) Then
                
'                   Si la décision est favorable, ou s'il s'agit d'une décision non déterminée du Conseil d'État suite à une décision favorable de la Cour Administrative d'Appel
                    If plngDécision = adpACCORD Or (![Provenance] = gclsRecoursContentieux.lngConseilÉtat And IsNull(pvarDateRetrait) And plngDécision = AdoptionsDélibérations.adpINDÉTERMINÉ) Then
'                       Suspendre le retrait
                        pvarDateRetrait = Null
'                   Sinon
                    Else
'                       Valider le retrait
                        pvarDateRetrait = pvarDateRetraitInitiale
                    End If
                    
                End If
                
            End If

            .MoveNext
            
        Wend
        
    End With
    
'   Mettre à jour les données relatives à l'agrément
    modDAO.SQLMettreÀJour daoMISE_À_JOUR_002, True, pstrMessageSiVerrouillage, plngNuméroDemande, pvarAgrémentDate, pvarAgrémentProvenance, pvarDateRetrait
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then Exit Sub
                            
    Exit Sub

Err_ActualiserDonnéesAgrément:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Actualiser les données relatives aux intervenants présents sur une demande
'   plngNuméroDemande:          numéro de la demande
'   pstrMessageSiVerrouillage:  message à afficher si les données sont verrouillées par une transaction

Public Sub ActualiserDonnéesIntervenants(plngNuméroDemande As Long, pstrMessageSiVerrouillage As String)
Const pkstrNomProcédure As String = "ActualiserDonnéesIntervenants"
Dim pvaru01_booHistoIntervenantPsychologue As Variant
Dim pvarTRSO_NUMPSY As Variant, pvarDEMA_DDEMRP As Variant, pvarDEMA_DPRRP As Variant, pvarDEMA_DRECRP As Variant
Dim pvarTRSO_NUMAS As Variant, pvarDEMA_DDEMRA As Variant, pvarDEMA_DPRRA As Variant, pvarDEMA_DRECRA
Dim pvarTRSO_NUMPSY_Initial As Variant, pvarDEMA_DDEMRP_Initial As Variant, pvarDEMA_DPRRP_Initial As Variant, pvarDEMA_DRECRP_Initial As Variant
Dim pvarTRSO_NUMAS_Initial As Variant, pvarDEMA_DDEMRA_Initial As Variant, pvarDEMA_DPRRA_Initial As Variant, pvarDEMA_DRECRA_Initial As Variant
Dim prstHistoriqueIntervenants As Recordset
On Error GoTo Err_ActualiserDonnéesIntervenants

'   Lire l'historique des intervenants
    Set prstHistoriqueIntervenants = modDAO.rstSQLExtraction(daoEXTRACTION_024, True, pstrMessageSiVerrouillage, plngNuméroDemande)
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub
    
    pvaru01_booHistoIntervenantPsychologue = Null
    pvarTRSO_NUMPSY = Null: pvarDEMA_DDEMRP = Null: pvarDEMA_DPRRP = Null: pvarDEMA_DRECRP = Null
    pvarTRSO_NUMAS = Null: pvarDEMA_DDEMRA = Null: pvarDEMA_DPRRA = Null: pvarDEMA_DRECRA = Null
    pvarTRSO_NUMPSY_Initial = Null: pvarDEMA_DDEMRP_Initial = Null: pvarDEMA_DPRRP_Initial = Null: pvarDEMA_DRECRP_Initial = Null
    pvarTRSO_NUMAS_Initial = Null: pvarDEMA_DDEMRA_Initial = Null: pvarDEMA_DPRRA_Initial = Null: pvarDEMA_DRECRA_Initial = Null
    
    With prstHistoriqueIntervenants
    
'       Pour chaque intervenant
        While Not .EOF
        
'           S'il s'agit du psychologue ou de l'assistante sociale en charge
            If IsNull(pvaru01_booHistoIntervenantPsychologue) Or Not pvaru01_booHistoIntervenantPsychologue = ![u01_booHistoIntervenantPsychologue] Then
            
                pvaru01_booHistoIntervenantPsychologue = ![u01_booHistoIntervenantPsychologue]
                
                Select Case ![u01_booHistoIntervenantPsychologue]
'                   Psychologue
                    Case True:  pvarTRSO_NUMPSY = ![TRSO_NUM]:          pvarDEMA_DDEMRP = ![u01_dtmHistoIntervenantDemande]:            pvarDEMA_DPRRP = ![u01_dtmHistoIntervenantPrevu]:           pvarDEMA_DRECRP = ![u01_dtmHistoIntervenantRecu]
'                   Assistante sociale
                    Case False: pvarTRSO_NUMAS = ![TRSO_NUM]:           pvarDEMA_DDEMRA = ![u01_dtmHistoIntervenantDemande]:            pvarDEMA_DPRRA = ![u01_dtmHistoIntervenantPrevu]:           pvarDEMA_DRECRA = ![u01_dtmHistoIntervenantRecu]
                End Select
            
            End If
            
'           Si l'historique se rapporte à la demande de rapport initiale
            If ![u01_intHistoIntervenantCadre] = AdoptionsCadresInterventions.adpDEMANDE_INITIALE Then
            
                Select Case ![u01_booHistoIntervenantPsychologue]
'                   Psychologue
                    Case True:  pvarTRSO_NUMPSY_Initial = ![TRSO_NUM]:  pvarDEMA_DDEMRP_Initial = ![u01_dtmHistoIntervenantDemande]:    pvarDEMA_DPRRP_Initial = ![u01_dtmHistoIntervenantPrevu]:   pvarDEMA_DRECRP_Initial = ![u01_dtmHistoIntervenantRecu]
'                   Assistante sociale
                    Case False: pvarTRSO_NUMAS_Initial = ![TRSO_NUM]:   pvarDEMA_DDEMRA_Initial = ![u01_dtmHistoIntervenantDemande]:    pvarDEMA_DPRRA_Initial = ![u01_dtmHistoIntervenantPrevu]:   pvarDEMA_DRECRA_Initial = ![u01_dtmHistoIntervenantRecu]
                End Select
            
            End If
            
            .MoveNext
            
        Wend
        
    End With

'   Mettre à jour les données relatives aux intervenants
    modDAO.SQLMettreÀJour daoMISE_À_JOUR_004, True, pstrMessageSiVerrouillage, plngNuméroDemande, pvarTRSO_NUMAS, pvarTRSO_NUMPSY, pvarDEMA_DDEMRP, pvarDEMA_DPRRP, pvarDEMA_DRECRP, pvarDEMA_DDEMRA, pvarDEMA_DPRRA, pvarDEMA_DRECRA, pvarTRSO_NUMPSY_Initial, pvarDEMA_DDEMRP_Initial, pvarDEMA_DPRRP_Initial, pvarDEMA_DRECRP_Initial, pvarTRSO_NUMAS_Initial, pvarDEMA_DDEMRA_Initial, pvarDEMA_DPRRA_Initial, pvarDEMA_DRECRA_Initial
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then Exit Sub
                            
    Exit Sub

Err_ActualiserDonnéesIntervenants:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Afficher l'identité du demandeur
'   pstrNuméroDossier:          numéro de dossier
'   ptxtNom:                    contrôle affichant l'identité du demandeur
'   pstrMessageSiVerrouillage:  message à afficher si les données sont verrouillées par une transaction
'   pcmdÀDésactiver:            CommandButton à désactiver dans le cas où les données sont verrouillées par une transaction [facultatif]

Public Sub AfficherIdentitéDemandeur(pstrNuméroDossier As String, ptxtNom As textbox, pstrMessageSiVerrouillage As String, Optional pcmdÀDésactiver As CommandButton)
Const pkstrNomProcédure As String = "AfficherIdentitéDemandeur"
Dim prstDemandeur As Recordset
On Error GoTo Err_AfficherIdentitéDemandeur

'   Lire l'identité du demandeur
    Set prstDemandeur = modDAO.rstSQLExtraction(daoINDIVIDU, True, pstrMessageSiVerrouillage, pstrNuméroDossier, gstrTYPE_DEMANDEUR)
    Select Case modDAO.lngDernièreErreurDAO()
'       Les données sont verrouillées par une transaction
        Case daoENREGISTREMENT_VERROUILLÉ:  If Not pcmdÀDésactiver Is Nothing Then modCommun.DéfinirVerrouillageContrôle pcmdÀDésactiver, True
'       Le jeu ne contient pas d'enregistrement
        Case daoAUCUN_ENREGISTREMENT:       Err.Raise mesDEMANDEUR_NON_DÉCLARÉ
    End Select
    
'   L'afficher
    If modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
        With prstDemandeur
            ptxtNom = Trim(Trim(![INDI_NOM]) & " " & Trim(![INDI_PREN]))
        End With
    End If
                            
    Exit Sub

Err_AfficherIdentitéDemandeur:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:             pintTraitement = mesINTERROMPRE
        Case mesDEMANDEUR_NON_DÉCLARÉ:  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
        Case Else:                      pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Afficher l'identité de l'intervenant
'   pintIntervenantNuméro:      numéro de l'intervenant
'   ptxtIntervenantNom:         contrôle affichant l'identité de l'intervenant
'   pstrMessageSiVerrouillage:  message à afficher si les données sont verrouillées par une transaction
'   pcmdÀDésactiver:            CommandButton à désactiver dans le cas où les données sont verrouillées par une transaction [facultatif]

Public Sub AfficherIdentitéIntervenant(pintIntervenantNuméro As Integer, ptxtIntervenantNom As textbox, pstrMessageSiVerrouillage As String, Optional pcmdÀDésactiver As CommandButton)
Const pkstrNomProcédure As String = "AfficherIdentitéIntervenant"
Dim prstIntervenant As Recordset
On Error GoTo Err_AfficherIdentitéIntervenant

'   Lire le nom de l'intervenant
    ptxtIntervenantNom = txtNomIntervenant(pintIntervenantNuméro, pstrMessageSiVerrouillage)
    If modDAO.lngDernièreErreurDAO() = daoENREGISTREMENT_VERROUILLÉ Then
        If Not pcmdÀDésactiver Is Nothing Then modCommun.DéfinirVerrouillageContrôle pcmdÀDésactiver, True
    End If
                            
    Exit Sub

Err_AfficherIdentitéIntervenant:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Retourner un booléen précisant si la saisie ou la suppression d'une décision est incohérente
'   plngNuméroDemande:          numéro de la demande
'   pstrMessageSiVerrouillage:  message à afficher si les données sont verrouillées par une transaction
'   plngTraitement:             type de traitement
'   pvarCommissionRéférence:    date de réunion de la commission à l'origine des recours éventuels
'   plngProvenance:             provenance de la décision
'   pdtpDécision:               DTPicker permettant la saisie:
'                               - de la date de la commission   (s'il s'agit d'une commission initiale ou d'un recours administratif)
'                               - de la décision                (s'il s'agit d'un recours contentieux)
'   pcboDécision:               ComboBox permettant la saisie de la décision
'   pdtpSignature:              DTPicker permettant la saisie:
'                               - de la date de signature       (s'il s'agit d'une commission initiale ou d'un recours administratif)
'                               - de la date d'agrément         (s'il s'agit d'un recours contentieux)
'   pdtpRecoursDépôt:           DTPicker permettant la saisie de la date de dépôt du recours                                        [obligatoire s'il s'agit d'un recours]
'   pdtmRefusRetraitSignature:  date de signature de la décision de refus ou de retrait à l'origine du recours                      [obligatoire s'il s'agit d'un recours]
'   pstrRefusRetrait:           libellé de la décision (<refus> ou <retrait>)                                                       [obligatoire s'il s'agit d'un recours]
'   pdtpDécisionPrévue:         DTPicker permettant la saisie de la date prévue pour la décision                                    [facultatif]
'   pcmdÀDésactiver:            CommandButton à désactiver selon l'erreur                                                           [facultatif]

Public Function booSaisieIncohérenteDécision(plngNuméroDemande As Long, pstrMessageSiVerrouillage As String, plngTraitement As CommunTraitements, pvarCommissionRéférence As Variant, plngProvenance As AdoptionsProvenances, pdtpDécision As DTPicker, pcboDécision As ComboBox, pdtpSignature As DTPicker, Optional pdtpRecoursDépôt As DTPicker, Optional pdtmRefusRetraitSignature As Date, Optional pstrRefusRetrait As String, Optional pdtpDécisionPrévue As DTPicker, Optional pcmdÀDésactiver As CommandButton) As Boolean
Const pkstrNomProcédure As String = "booSaisieIncohérenteDécision"
Dim pvarCommissionPrécédente As Variant, pvarDépôt As Variant, pvarPremièreCommissionÀPrendreDécision As Variant, pvarDateSignatureDemandeInitiale As Variant, pvarDateAgrément As Variant, pvarDateRetrait As Variant, pvarDateDépôtPrécédent As Variant
Dim pctlFocus As Control
Dim pclsDécisions As clsDécisions, pclsDécision As clsDécision
Dim prstDécisions As Recordset
Dim plngDécisionCollection As AdoptionsDélibérations, plngDécisionCommissionInitiale As AdoptionsDélibérations, plngDécisionContentieuxPrécédente As AdoptionsDélibérations
Dim pdtmDécision As Date
Dim pstrMessages(1) As String, pstrTypeRecours As String, pstrTypeValidation As String
Dim plngErreur As Long
Dim pintItération As Integer
Dim pbooDécisionÀInsérer As Boolean, pbooRefuserValidation As Boolean
On Error GoTo Err_booSaisieIncohérenteDécision

    Select Case plngTraitement
        Case comCRÉATION:       pstrTypeValidation = "Création"
        Case comMODIFICATION:   pstrTypeValidation = "Mise à jour"
        Case comSUPPRESSION:    pstrTypeValidation = "Suppression"
    End Select
    
'   S'il ne s'agit pas d'une suppression
    If Not plngTraitement = comSUPPRESSION Then
    
'       Si la date prévue pour la décision doit être vérifiée
        If Not pdtpDécisionPrévue Is Nothing Then
        
'           Refuser la validation si la date prévue pour la décision n'est pas définie
            If IsNull(pdtpDécisionPrévue) Then
                Set pctlFocus = pdtpDécisionPrévue: pstrMessages(0) = "de la date prévue pour la décision"
                Err.Raise mesSAISIE_MANQUANTE
                booSaisieIncohérenteDécision = True: Exit Function
            End If
            
        End If
        
'       S'il s'agit d'un recours
        If Not pdtpRecoursDépôt Is Nothing Then
        
            pstrTypeRecours = IIf(plngProvenance = AdoptionsProvenances.adpRECOURS_ADMINISTRATIF, " administratif", "")
            
'           Refuser la validation si la date de dépôt n'est pas définie
            If IsNull(pdtpRecoursDépôt) Then
                Set pctlFocus = pdtpRecoursDépôt: pstrMessages(0) = "de la date de dépôt"
                Err.Raise mesSAISIE_MANQUANTE
                booSaisieIncohérenteDécision = True: Exit Function
            End If
        
'           Refuser la validation si la date de dépôt est antérieure à la date de signature du refus/retrait
            If pdtpRecoursDépôt < pdtmRefusRetraitSignature Then
                Set pctlFocus = pdtpRecoursDépôt: pstrMessages(0) = "de dépôt du recours" & pstrTypeRecours: pstrMessages(1) = "de signature du " & pstrRefusRetrait & " (" & Format(pdtmRefusRetraitSignature, gstrFORMAT_DATE_FRANÇAIS) & ")"
                Err.Raise mesDATE_ANTÉRIEURE
                booSaisieIncohérenteDécision = True: Exit Function
            End If

            pstrMessages(1) = "de dépôt du recours" & pstrTypeRecours
            
'           Refuser la validation si la date de la commission/décision est antérieure à la date de dépôt
            If pdtpDécision < pdtpRecoursDépôt Then
                Set pctlFocus = pdtpDécision: pstrMessages(0) = IIf(plngProvenance = AdoptionsProvenances.adpRECOURS_ADMINISTRATIF, "de la commission", "de la décision")
                Err.Raise mesDATE_ANTÉRIEURE
                booSaisieIncohérenteDécision = True: Exit Function
            End If
        
'           Refuser la validation si la date de la signature/de l'agrément est antérieure à la date de dépôt
            If pdtpSignature < pdtpRecoursDépôt Then
                Set pctlFocus = pdtpSignature: pstrMessages(0) = IIf(plngProvenance = AdoptionsProvenances.adpRECOURS_ADMINISTRATIF, "de la signature", "de l'agrément")
                Err.Raise mesDATE_ANTÉRIEURE
                booSaisieIncohérenteDécision = True: Exit Function
            End If
        
'           Si la date prévue pour la décision doit être vérifiée
            If Not pdtpDécisionPrévue Is Nothing Then
            
'               Refuser la validation si la date prévue pour la décision est antérieure à la date de dépôt
                If pdtpDécisionPrévue < pdtpRecoursDépôt Then
                    Set pctlFocus = pdtpDécisionPrévue: pstrMessages(0) = "prévue pour la décision"
                    Err.Raise mesDATE_ANTÉRIEURE
                    booSaisieIncohérenteDécision = True: Exit Function
                End If
                
            End If
            
        End If
        
        pstrMessages(1) = IIf(plngProvenance > adpINDÉTERMINÉE, "de la décision", "de la commission")
        
'       Refuser la validation si la date de la signature/de l'agrément est antérieure à la date de la commission/la décision
        If pdtpSignature < pdtpDécision Then
            Set pctlFocus = pdtpSignature: pstrMessages(0) = IIf(plngProvenance > adpINDÉTERMINÉE, "de l'agrément", "de la signature")
            Err.Raise mesDATE_ANTÉRIEURE
            booSaisieIncohérenteDécision = True: Exit Function
        End If
        
    End If

'   Lire l'ensemble des décisions relatives à une demande
'       Notes:
'           Voir les règles de gestion définies en entête de la feuille frm_DecisionCommission
'           Les décisions sont triées de façon ascendante par date de commission et provenance (Demande initiale, puis Recours administratif / Tribunal Administratif / Cour Administrative d'Appel / Conseil d'État)
'           Il existe 3 types d'enregistrements, chacun implémentant de façon différente les colonnes DateCommission et DateSignature:
'           - enregistrement issu de la saisie d'une commission:
'               DateCommission: date de la commission       (n'est jamais nulle)
'               DateSignature:  date de la signature        (peut être nulle)
'               DateDepot:      NULL
'           - enregistrement issu de la saisie d'un recours administratif:
'               DateCommission: date de la commission       (peut être nulle)
'               DateSignature:  date de la signature        (peut être nulle)
'                   dans le cas où la décision est prise par le Président du Conseil Général sans consultation d'une commission, la date de la commission est nulle et la date de la décision correspond à la date de la signature
'               DateDepot:      date du dépôt               (n'est jamais nulle)
'           - enregistrement issu de la saisie d'un recours contentieux:
'               DateCommission: NULL
'               DateSignature:  date de l'agrément          (peut être nulle)
'               DateDepot:      date du dépôt               (n'est jamais nulle)
    Set prstDécisions = modDAO.rstSQLExtraction(daoDÉCISIONS, True, pstrMessageSiVerrouillage, plngNuméroDemande)
    If modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT And Not plngTraitement = comCRÉATION Then
        Err.Raise mesMESSAGE_AVEC_ICÔNE
        If Not pcmdÀDésactiver Is Nothing Then modCommun.DéfinirVerrouillageContrôle pcmdÀDésactiver, True
    End If
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or (Not modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT Or plngTraitement = comCRÉATION)) Then booSaisieIncohérenteDécision = True: Exit Function
    
'   Initialiser la collection des décisions
    Set pclsDécisions = New clsDécisions
    
    With prstDécisions
    
'       Pour chaque décision
        While Not .EOF
        
            Select Case modCommun.varValeurFormatéeSiNull(![Decision], vbString)
                Case CStr(AdoptionsDélibérations.adpINDÉTERMINÉ), "":   plngDécisionCollection = AdoptionsDélibérations.adpINDÉTERMINÉ
                Case CStr(adpREFUS), gstrREFUS:                         plngDécisionCollection = adpREFUS
                Case CStr(adpRETRAIT):                                  plngDécisionCollection = adpRETRAIT
                Case CStr(adpACCORD), gstrACCORD:                       plngDécisionCollection = adpACCORD
            End Select
            
'           L'ajouter à la collection
            Set pclsDécision = pclsDécisions.Ajout(0, ![CommissionInitiale], ![Provenance], plngDécisionCollection, ![DateCommission], ![DateSignature], pvarDépôt:=![DateDepot], pvarClé:=Format(![CommissionInitiale], "yyyymmdd") & gstrSÉPARATEUR_1 & CStr(![Provenance] - AdoptionsProvenances.adpDEMANDE_INITIALE))
            
            .MoveNext
            
        Wend
        
    End With
    
'   S'il s'agit d'un suppression
    If plngTraitement = comSUPPRESSION Then
    
'       La supprimer de la collection
        pclsDécisions.Supprimer Format(CDate(pvarCommissionRéférence), "yyyymmdd") & gstrSÉPARATEUR_1 & CStr(plngProvenance - AdoptionsProvenances.adpDEMANDE_INITIALE)
        
'   Sinon
    Else
            
'       S'il s'agit de mettre à jour une décision, la supprimer de la collection
        If plngTraitement = comMODIFICATION Then pclsDécisions.Supprimer Format(CDate(pvarCommissionRéférence), "yyyymmdd") & gstrSÉPARATEUR_1 & CStr(plngProvenance - AdoptionsProvenances.adpDEMANDE_INITIALE)
    
        Select Case plngProvenance
            Case AdoptionsProvenances.adpDEMANDE_INITIALE:  pdtmDécision = pdtpDécision:                    pvarDépôt = Null
            Case Else:                                      pdtmDécision = CDate(pvarCommissionRéférence):  pvarDépôt = pdtpRecoursDépôt
        End Select
        
        pintItération = 0: plngErreur = 0
'       Pour chaque décision de la collection
        For Each pclsDécision In pclsDécisions
            pintItération = pintItération + 1
'           Si la décision précède celle de la collection
            If Format(pdtmDécision, "yyyymmdd") & gstrSÉPARATEUR_1 & CStr(plngProvenance - AdoptionsProvenances.adpDEMANDE_INITIALE) < pclsDécision.strClé Then
                pbooDécisionÀInsérer = True
                Exit For
            End If
        Next pclsDécision
'       Si la décision doit être insérée dans la collection
        If pbooDécisionÀInsérer Then
'           L'insérer
            pclsDécisions.Ajout plngErreur, pdtmDécision, plngProvenance, pcboDécision.ItemData(pcboDécision.ListIndex), IIf(plngProvenance < AdoptionsProvenances.adpINDÉTERMINÉE, pdtpDécision.Value, Null), pdtpSignature.Value, pvarDépôt:=pvarDépôt, pvarClé:=Format(pdtmDécision, "yyyymmdd") & gstrSÉPARATEUR_1 & CStr(plngProvenance - AdoptionsProvenances.adpDEMANDE_INITIALE), pvarAvant:=pintItération
'       Sinon
        Else
'           L'ajouter à la commission
            pclsDécisions.Ajout plngErreur, pdtmDécision, plngProvenance, pcboDécision.ItemData(pcboDécision.ListIndex), IIf(plngProvenance < AdoptionsProvenances.adpINDÉTERMINÉE, pdtpDécision.Value, Null), pdtpSignature.Value, pvarDépôt:=pvarDépôt, pvarClé:=Format(pdtmDécision, "yyyymmdd") & gstrSÉPARATEUR_1 & CStr(plngProvenance - AdoptionsProvenances.adpDEMANDE_INITIALE)
        End If
'       Refuser la validation si elle conduit à créer une doublon
        If plngErreur = comERREUR_00457 Then
            Set pctlFocus = pdtpDécision: pstrMessages(0) = IIf(plngProvenance = AdoptionsProvenances.adpDEMANDE_INITIALE, "La commission", "Le recours")
            Err.Raise mesEXISTE_DÉJÀ
            Set pclsDécisions = Nothing
            booSaisieIncohérenteDécision = True: Exit Function
        End If
            
    End If
            
    pvarCommissionPrécédente = Null: pvarPremièreCommissionÀPrendreDécision = Null: pvarDateAgrément = Null:  pvarDateRetrait = Null
    
'   Pour chaque décision de la collection
    For Each pclsDécision In pclsDécisions
    
        With pclsDécision
        
'           Si la décision fait référence à une nouvelle commission initiale
            If IsNull(pvarCommissionPrécédente) Or Not pvarCommissionPrécédente = .dtmCommissionInitiale Then
                pvarCommissionPrécédente = .dtmCommissionInitiale
                plngDécisionCommissionInitiale = AdoptionsDélibérations.adpINDÉTERMINÉ
                pvarDateSignatureDemandeInitiale = Null: pvarDateDépôtPrécédent = Null
                pintItération = gclsRecoursContentieux.clsRecoursContentieux(1).lngCode - 1
            End If
            
'           S'il s'agit d'une commission initiale
            If .lngProvenance = AdoptionsProvenances.adpDEMANDE_INITIALE Then
            
                plngDécisionCommissionInitiale = .lngDécision
                pvarDateSignatureDemandeInitiale = .varSignature
                
'               Si la première commission à prendre une décision n'a pas encore été identifiée
                If IsNull(pvarPremièreCommissionÀPrendreDécision) Then
                
'                   Si la décision est définie, référencer la commission comme la première à prendre une décision
                    If Not .lngDécision = AdoptionsDélibérations.adpINDÉTERMINÉ Then pvarPremièreCommissionÀPrendreDécision = .dtmCommissionInitiale
                    
'               Sinon
                Else
                
'                   Refuser la validation si aucun agrément n'a été délivré
                    If IsNull(pvarDateAgrément) Then
                        pstrMessages(0) = "L'absence d'agrément":                                                                           pstrMessages(1) = Format(.dtmCommissionInitiale, gstrFORMAT_DATE_FRANÇAIS)
                        pbooRefuserValidation = True
'                   Sinon, refuser la validation si la date de commission est antérieure à la délivrance de l'agrément
                    ElseIf .dtmCommissionInitiale < pvarDateAgrément Then
                        pstrMessages(0) = "L'agrément, délivré le " & Format(pvarDateAgrément, gstrFORMAT_DATE_FRANÇAIS) & " seulement,":    pstrMessages(1) = Format(.dtmCommissionInitiale, gstrFORMAT_DATE_FRANÇAIS)
                        pbooRefuserValidation = True
'                   Sinon, refuser la validation si le retrait a été décidé
                    ElseIf Not IsNull(pvarDateRetrait) Then
                        pstrMessages(0) = "Le retrait décidé le " & Format(pvarDateRetrait, gstrFORMAT_DATE_FRANÇAIS):                      pstrMessages(1) = Format(.dtmCommissionInitiale, gstrFORMAT_DATE_FRANÇAIS)
                        pbooRefuserValidation = True
                    End If
                    
                    If pbooRefuserValidation Then
                        Err.Raise mesMESSAGE_016
                        Set pclsDécisions = Nothing
                        booSaisieIncohérenteDécision = True: Exit Function
                    End If
                    
                End If

'           Sinon
            Else
            
                pstrTypeRecours = IIf(.lngProvenance = AdoptionsProvenances.adpRECOURS_ADMINISTRATIF, "du recours administratif", "d'un recours contentieux")
            
'               Refuser la validation s'il s'agit d'un recours administratif ou du premier recours contentieux, et si le recours ne fait suite à aucune décision de refus ou de retrait
                If (.lngProvenance = AdoptionsProvenances.adpRECOURS_ADMINISTRATIF Or .lngProvenance = AdoptionsProvenances.adpINDÉTERMINÉE + 1) And Not (plngDécisionCommissionInitiale = adpREFUS Or plngDécisionCommissionInitiale = adpRETRAIT) Then
                    pstrMessages(0) = IIf(.lngProvenance = AdoptionsProvenances.adpRECOURS_ADMINISTRATIF, "administratif", "contentieux")
                    If Not IsNull(.varDépôt) Then pstrMessages(1) = " déposé le " & Format(.varDépôt, gstrFORMAT_DATE_FRANÇAIS)
                    Err.Raise mesMESSAGE_015
                    Set pclsDécisions = Nothing
                    booSaisieIncohérenteDécision = True: Exit Function
                End If

'               Refuser la validation si la date de signature de la décision initiale n'est pas saisie
                If IsNull(pvarDateSignatureDemandeInitiale) Then
                    pstrMessages(0) = LCase(gclsDélibérations(CStr(plngDécisionCommissionInitiale)).strLibellé): pstrMessages(1) = Format(.varDépôt, gstrFORMAT_DATE_FRANÇAIS)
                    Err.Raise mesMESSAGE_019
                    Set pclsDécisions = Nothing
                    booSaisieIncohérenteDécision = True: Exit Function
                End If

'               Refuser la validation si la date de dépôt est antérieure à la date de signature de la décision initiale
                If .varDépôt < pvarDateSignatureDemandeInitiale Then
                    pstrMessages(0) = "de dépôt " & pstrTypeRecours & " (" & Format(.varDépôt, gstrFORMAT_DATE_FRANÇAIS) & ")": pstrMessages(1) = "de signature du " & LCase(gclsDélibérations(CStr(plngDécisionCommissionInitiale)).strLibellé) & " (" & Format(pvarDateSignatureDemandeInitiale, gstrFORMAT_DATE_FRANÇAIS) & ")"
                    Err.Raise mesDATE_ANTÉRIEURE
                    booSaisieIncohérenteDécision = True: Exit Function
                End If
        
'               S'il s'agit d'un recours contentieux
                If .lngProvenance > AdoptionsProvenances.adpINDÉTERMINÉE Then
                
                    pintItération = pintItération + 1
                    
'                   Refuser la validation si la succession des recours ne respecte pas l'ordre imposé par le paramétrage
                    If Not .lngProvenance = pintItération Then
                        Err.Raise mesMESSAGE_017
                        If Not pcmdÀDésactiver Is Nothing Then modCommun.DéfinirVerrouillageContrôle pcmdÀDésactiver, True
                        Set pclsDécisions = Nothing
                        booSaisieIncohérenteDécision = True: Exit Function
                    End If
                    
'                   S'il ne s'agit pas du premier recours
                    If Not IsNull(pvarDateDépôtPrécédent) Then
                    
'                       Refuser la validation si le recours précédent n'a pas conduit à une décision
                        If plngDécisionContentieuxPrécédente = AdoptionsDélibérations.adpINDÉTERMINÉ Then
                            pstrMessages(0) = modCommun.varValeurFormatéeSiNull(.varDépôt, vbDate, "<date inconnue>"): pstrMessages(1) = modCommun.varValeurFormatéeSiNull(pvarDateDépôtPrécédent, vbDate, "<date inconnue>")
                            Err.Raise mesMESSAGE_018
                            Set pclsDécisions = Nothing
                            booSaisieIncohérenteDécision = True: Exit Function
                        End If

'                       Refuser la validation si la date de dépôt est antérieure à la date de dépôt du recours précédent
                        If .varDépôt < pvarDateDépôtPrécédent Then
                            Set pctlFocus = pdtpRecoursDépôt
                            Err.Raise mesMESSAGE_014
                            Set pclsDécisions = Nothing
                            booSaisieIncohérenteDécision = True: Exit Function
                        End If
                        
                    End If
                    
                    pvarDateDépôtPrécédent = .varDépôt
                    plngDécisionContentieuxPrécédente = .lngDécision
                    
'               Sinon
                Else
                    pvarDateDépôtPrécédent = Null
                End If

            End If
                    
            Select Case .lngDécision
            
'               Refus
                Case adpREFUS
                    
                    pvarDateAgrément = Null
                    
'               Retrait
                Case adpRETRAIT
                    
'                   Si le retrait ne fait suite à aucun agrément
                    If IsNull(pvarDateAgrément) Then
                        pstrMessages(0) = CStr(pclsDécision.dtmCommissionInitiale)
                        Err.Raise mesMESSAGE_007
                        pbooRefuserValidation = True
'                   Sinon
                    Else
                        pvarDateRetrait = .dtmCommissionInitiale
                    End If

'               Accord
                Case adpACCORD
                
'                   Si la décision fait référence à la première commission à prendre une décision
                    If .dtmCommissionInitiale = pvarPremièreCommissionÀPrendreDécision Then
'                       S'il ne s'agit pas d'une décision du Conseil d'État ou si la décision précédente n'était pas favorable
                        If Not .lngProvenance = gclsRecoursContentieux.lngConseilÉtat Or IsNull(pvarDateAgrément) Then
'                           Faire de ...
'                               - la date de réunion de la commission (cas de la commission initiale)
'                               - la date de réunion de la commission, ou la date de signature si la décision est prise par le Président du Conseil Général sans consultation d'une commission (cas du recours administratif)
'                               - la date de délivrance de l'agrément par le Président du Conseil Général (cas du recours contentieux)
'                           ... la date de délivrance de l'agrément initial
                            pvarDateAgrément = IIf(IsNull(.varCommission), .varSignature, .varCommission)
                        End If
                    End If
                    
            End Select
            
        End With
        
'       Si la cohérence de la décision n'est pas vérifiée
        If pbooRefuserValidation Then
            Set pclsDécisions = Nothing
            booSaisieIncohérenteDécision = True: Exit Function
        End If
                
    Next pclsDécision
    
    Set pclsDécisions = Nothing

    Exit Function
    
Err_booSaisieIncohérenteDécision:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:         pintTraitement = mesINTERROMPRE
        Case mesMESSAGE_AVEC_ICÔNE: pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrTypeValidation & " refusée: " & gstrDONNÉES_SUPPRIMÉES, vbExclamation)
        Case mesSAISIE_MANQUANTE:   pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, pctlFocus, pstrMessages(0))
        Case mesDATE_ANTÉRIEURE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, pctlFocus, pstrMessages(0), pstrMessages(1))
        Case mesEXISTE_DÉJÀ:        pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, pctlFocus, pstrTypeValidation, pstrMessages(0))
        Case mesMESSAGE_007:        pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrTypeValidation, pstrMessages(0))
        Case mesMESSAGE_014:        pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, pctlFocus, pstrTypeValidation)
        Case mesMESSAGE_015:        pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrTypeValidation, pstrMessages(0), pstrMessages(1))
        Case mesMESSAGE_016:        pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrTypeValidation, pstrMessages(0), pstrMessages(1))
        Case mesMESSAGE_017:        pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrTypeValidation)
        Case mesMESSAGE_018:        pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrTypeValidation, pstrMessages(0), pstrMessages(1))
        Case mesMESSAGE_019:        pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrTypeValidation, pstrMessages(0), pstrMessages(1))
        Case Else:                  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Function


'Retourner un booléen précisant si la saisie d'un intervenant est incohérente
'   pdtpRéférenceCalculRéceptionPrévue:         DTPicker permettant la saisie de la date de référence pour le calcul de la date prévue de réception
'   pdtmRéférencePourDemande:                   date de référence pour la demande [Exemple: date de dépôt du recours pour un recours administratif, date de réception du dossier pour une demande initiale ou une surveillance de placement]
'   pvarIntervenantNuméro:                      numéro de l'intervenant (Null si aucun intervenant n'est sélectionné)
'   pcmdIntervenant:                            CommandButton permettant la sélection de l'intervenant
'   pdtpDemande:                                DTPicker permettant la saisie de la date de la demande (du rapport ou de la surveillance)
'   pdtpReçu:                                   DTPicker permettant la saisie de la date de réception du rapport
'   pstrChaîneIntervenant:                      chaîne utilisée dans la construction du message en cas de saisie incohérente [Exemple: "du psychologue"]
'   pstrChaîneDateRéférenceCalculRéception:     chaîne utilisée dans la construction du message en cas de saisie incohérente [Exemple: "de référence pour le calcul de l'échéancier"]
'   pstrChaîneDateRéférencePourDemandeRapport:  chaîne utilisée dans la construction du message en cas de saisie incohérente [Exemple: "de demande du dossier"]
'   pdtpPrévu:                                  DTPicker permettant la saisie de la date prévue pour la réception du rapport [obligatoire si pdtpIntervention n'est pas passé en argument]
'   pdtpIntervention:                           DTPicker permettant la saisie de la date d'intervention                      [obligatoire si pdtpPrévu n'est pas passé en argument]

Public Function booSaisieIncohérenteIntervenant(pdtpRéférenceCalculRéceptionPrévue As DTPicker, pdtmRéférencePourDemande As Date, pvarIntervenantNuméro As Variant, pcmdIntervenant As CommandButton, pdtpDemande As DTPicker, pdtpReçu As DTPicker, pstrChaîneIntervenant As String, pstrChaîneDateRéférenceCalculRéception As String, pstrChaîneDateRéférencePourDemandeRapport As String, Optional pdtpPrévu As DTPicker, Optional pdtpIntervention As DTPicker) As Boolean
Const pkstrNomProcédure As String = "booSaisieIncohérenteIntervenant"
Dim pvarDatePrévuOuIntervention As Variant
Dim pctlFocus As Control
Dim pstrMessages(1) As String
Dim pbooVérifierRéceptionPrévue As Boolean
On Error GoTo Err_booSaisieIncohérenteIntervenant

    pbooVérifierRéceptionPrévue = Not pdtpPrévu Is Nothing
    If pbooVérifierRéceptionPrévue Then pvarDatePrévuOuIntervention = pdtpPrévu.Value Else pvarDatePrévuOuIntervention = pdtpIntervention.Value
    
'   Refuser la validation si l'identité de l'intervenant est définie mais pas la date de la demande
    If Not IsNull(pvarIntervenantNuméro) And IsNull(pdtpDemande) Then
        Set pctlFocus = pdtpDemande: pstrMessages(0) = "La sélection " & pstrChaîneIntervenant: pstrMessages(1) = "à saisir une date pour la demande " & IIf(pbooVérifierRéceptionPrévue, "du rapport", "de la surveillance")
        Err.Raise mesOBLIGE
        booSaisieIncohérenteIntervenant = True
        Exit Function
    End If

'   Refuser la validation si l'une des dates relatives à l'intervenant est saisie mais pas son identité
    If IsNull(pvarIntervenantNuméro) And Not (IsNull(pdtpDemande) And IsNull(pdtpReçu) And IsNull(pvarDatePrévuOuIntervention)) Then
        Set pctlFocus = pcmdIntervenant: pstrMessages(0) = "La saisie d'une date liée au rapport " & pstrChaîneIntervenant: pstrMessages(1) = "à l'identifier"
        Err.Raise mesOBLIGE
        booSaisieIncohérenteIntervenant = True
        Exit Function
    End If

'   Si la date de la demande est définie
    If Not IsNull(pdtpDemande) Then
    
'       Refuser la validation si la date de la demande est antérieure à la date de référence
        If pdtpDemande < modCommun.varValeurFormatéeSiNull(pdtmRéférencePourDemande, vbDate, #12:00:00 AM#) Then
            Set pctlFocus = pdtpDemande: pstrMessages(0) = "de demande " & IIf(pbooVérifierRéceptionPrévue, "du rapport " & pstrChaîneIntervenant & vbCrLf, "de la surveillance"): pstrMessages(1) = pstrChaîneDateRéférencePourDemandeRapport & " (" & Format(pdtmRéférencePourDemande, gstrFORMAT_DATE_FRANÇAIS) & ")"
            Err.Raise mesDATE_ANTÉRIEURE
            booSaisieIncohérenteIntervenant = True
            Exit Function
        End If

'       Refuser la validation si la date de réception du rapport de l'intervenant est antérieure à la date de la demande
        If modCommun.varValeurFormatéeSiNull(pdtpReçu, vbDate) < pdtpDemande Then
            Set pctlFocus = pdtpReçu: pstrMessages(0) = "de réception du rapport " & pstrChaîneIntervenant & vbCrLf: pstrMessages(1) = IIf(pbooVérifierRéceptionPrévue, "de sa demande", "de la demande de surveillance")
            Err.Raise mesDATE_ANTÉRIEURE
            booSaisieIncohérenteIntervenant = True
            Exit Function
        End If

'       Si la date prévue pour la réception du rapport doit être vérifiée
        If pbooVérifierRéceptionPrévue Then
        
'           Refuser la validation si la date prévue pour la réception du rapport de l'intervenant est antérieure à sa demande
            If modCommun.varValeurFormatéeSiNull(pdtpPrévu, vbDate) < pdtpDemande Then
                Set pctlFocus = pdtpPrévu: pstrMessages(0) = "prévue pour la réception du rapport " & pstrChaîneIntervenant & vbCrLf: pstrMessages(1) = "de sa demande"
                Err.Raise mesDATE_ANTÉRIEURE
                booSaisieIncohérenteIntervenant = True
                Exit Function
            End If
    
'           Refuser la validation si la date prévue pour la réception du rapport de l'intervenant est antérieure à sa date de référence
            If modCommun.varValeurFormatéeSiNull(pdtpPrévu, vbDate) < modCommun.varValeurFormatéeSiNull(pdtpRéférenceCalculRéceptionPrévue, vbDate, #12:00:00 AM#) Then
                Set pctlFocus = pdtpPrévu: pstrMessages(0) = "prévue pour la réception du rapport " & pstrChaîneIntervenant: pstrMessages(1) = pstrChaîneDateRéférenceCalculRéception & " (" & Format(pdtpRéférenceCalculRéceptionPrévue, gstrFORMAT_DATE_FRANÇAIS) & ")"
                Err.Raise mesDATE_ANTÉRIEURE
                booSaisieIncohérenteIntervenant = True
                Exit Function
            End If

'       Sinon
        Else
        
'           Refuser la validation si la date d'intervention est antérieure à la demande de surveillance
            If modCommun.varValeurFormatéeSiNull(pdtpIntervention, vbDate) < pdtpDemande Then
                Set pctlFocus = pdtpIntervention: pstrMessages(0) = "de l'intervention " & pstrChaîneIntervenant & vbCrLf: pstrMessages(1) = "de la demande de surveillance"
                Err.Raise mesDATE_ANTÉRIEURE
                booSaisieIncohérenteIntervenant = True
                Exit Function
            End If
    
        End If
        
    End If

    Exit Function
    
Err_booSaisieIncohérenteIntervenant:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:         pintTraitement = mesINTERROMPRE
        Case mesDATE_ANTÉRIEURE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, pctlFocus, pstrMessages(0), pstrMessages(1))
        Case mesOBLIGE:             pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, pctlFocus, pstrMessages(0), pstrMessages(1))
        Case Else:                  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Function


'Centrer la feuille
'   pfrmCible:  feuille dont on souhaite le centrage

Public Sub CentrerFeuille(pfrmCible As Form)
Const pkstrNomProcédure As String = "CentrerFeuille"
Dim psngTop As Single
On Error GoTo Err_CentrerFeuille

    With pfrmCible
        
'       Calculer la position du haut de la fenêtre
        psngTop = adpMENU_HAUTEUR + ((((Screen.Height - adpDEMANDES_À_RÉACTUALISER_HAUTEUR) - adpDEMANDES_À_RÉACTUALISER_DÉCALAGE_BAS) - adpMENU_HAUTEUR - .Height) / 2)
        
'       La centrer
        .Move (Screen.Width - .Width) / 2, IIf(psngTop > 0, psngTop, (Screen.Height - .Height) / 2)
        
    End With
                            
    Exit Sub

Err_CentrerFeuille:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Implémenter les collections
'   pvarArgument(): tableau de type ParamArray constitué de la liste des collections à implémenter

Public Sub ImplémenterCollections(ParamArray pvarArgument() As Variant)
Const pkstrNomProcédure As String = "ImplémenterCollections"
Dim prstCollection As Recordset
Dim plngCode As Long
Dim pintIndiceArgument As Integer
On Error GoTo Err_ImplémenterCollections

'   Pour chaque argument passé en référence
    For pintIndiceArgument = LBound(pvarArgument) To UBound(pvarArgument)
    
        Select Case pvarArgument(pintIndiceArgument)
            
'           Délibérations
            Case "Délibérations"

                Set gclsDélibérations = Nothing: Set gclsDélibérations = New clsCodesLibellés
                gclsDélibérations.Ajout AdoptionsDélibérations.adpINDÉTERMINÉ, "", pvarClé:=CStr(AdoptionsDélibérations.adpINDÉTERMINÉ)
                gclsDélibérations.Ajout adpREFUS, "Refus", pvarClé:=CStr(adpREFUS)
                gclsDélibérations.Ajout adpRETRAIT, "Retrait", pvarClé:=CStr(adpRETRAIT)
                gclsDélibérations.Ajout adpACCORD, "Accord", pvarClé:=CStr(adpACCORD)
                
'           Cadres d'intervention
            Case "Cadres d'intervention"

                Set gclsCadresInterventions = Nothing: Set gclsCadresInterventions = New clsCodesLibellés
                gclsCadresInterventions.Ajout AdoptionsCadresInterventions.adpINDÉTERMINÉ, "<Indéterminé>", pvarClé:=CStr(AdoptionsCadresInterventions.adpINDÉTERMINÉ)
                gclsCadresInterventions.Ajout AdoptionsCadresInterventions.adpDEMANDE_INITIALE, "Demande initiale", "une ", pvarClé:=CStr(AdoptionsCadresInterventions.adpDEMANDE_INITIALE)
                gclsCadresInterventions.Ajout adpSECONDE_ÉVALUATION, "Seconde évaluation", "une ", pvarClé:=CStr(adpSECONDE_ÉVALUATION)
                gclsCadresInterventions.Ajout adpREPRISE, "Reprise de l'instruction", pvarClé:=CStr(adpREPRISE)
                gclsCadresInterventions.Ajout AdoptionsCadresInterventions.adpRECOURS_ADMINISTRATIF, "Recours administratif", "un ", pvarClé:=CStr(AdoptionsCadresInterventions.adpRECOURS_ADMINISTRATIF)
                gclsCadresInterventions.Ajout adpSURVEILLANCE_PLACEMENT, "Surveillance de placement", "une ", pvarClé:=CStr(adpSURVEILLANCE_PLACEMENT)
                
        End Select
    
    Next pintIndiceArgument
    
    Exit Sub
    
Err_ImplémenterCollections:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub


'Implémenter un contrôle ComboBox ou ListBox
'   pobjCible:          contrôle dont on souhaite l'implémentation
'   pstrType:           type d'implémentation
'   pstrListeMessages:  liste de messages déterminant les messages supplémentaires à afficher en cas d'erreur (voir modDAO.rstSQLExtraction)
'   pctlÀDésactiver:    contrôle à désactiver dans le cas où les données sont verrouillées par une transaction
'   pvarArgument():     tableau d'arguments
'       Cadres d'interventions
'           pvarArgument(0):    booléen précisant si les valeurs <Recours administratif> et <Surveillance de placement> doivent être affichées ou non
'       Délibérations
'           pvarArgument(0):    booléen précisant si la valeur <Retrait> doit être affichée ou non
'       Demandes
'           pvarArgument(0):    numéro de dossier
'       Sexes
'           pvarArgument(0):    booléen précisant si la valeur <Indéterminé> doit être affichée ou non
'           pvarArgument(1):    chaîne au format '<valeur Indéterminé>,<valeur Masculin>,<valeur Féminin>' ou '<valeur Masculin>,<valeur Féminin>' précisant les valeurs affichées [facultatif: si elle est manquante, la chaîne sera 'Indéterminé,Masculin,Féminin' ou 'Masculin,Féminin']

Public Sub ImplémenterListe(pobjCible As Object, pstrType As String, pstrListeMessages As String, pctlÀDésactiver As Control, ParamArray pvarArgument() As Variant)
Const pkstrNomProcédure As String = "ImplémenterListe"
Dim pclsCodeLibellé As clsCodeLibellé
Dim prstJeuEnregistrements As Recordset
Dim pstrValeurs() As String
On Error GoTo Err_ImplémenterListe

    pobjCible.Clear
    
    Select Case pstrType
    
'       Cadres d'interventions
        Case "Cadres d'interventions"
        
            modCommun.ImplémenterListe pobjCible, "", pvarNuméro:=adpAUCUNE_SÉLECTION
            
            For Each pclsCodeLibellé In gclsCadresInterventions
                With pclsCodeLibellé
                    If Not (.lngCode = AdoptionsCadresInterventions.adpRECOURS_ADMINISTRATIF Or .lngCode = adpSURVEILLANCE_PLACEMENT) Or pvarArgument(0) Then modCommun.ImplémenterListe pobjCible, .strLibellé, pvarNuméro:=.lngCode
                End With
            Next pclsCodeLibellé
    
'       Délibérations
        Case "Délibérations"
        
            For Each pclsCodeLibellé In gclsDélibérations
                With pclsCodeLibellé
                    If Not .lngCode = adpRETRAIT Or pvarArgument(0) Then modCommun.ImplémenterListe pobjCible, .strLibellé, pvarNuméro:=.lngCode
                End With
            Next pclsCodeLibellé
    
'       Numéros de demandes
        Case "Demandes"
        
'           Lire les numéros de demandes
            Set prstJeuEnregistrements = modDAO.rstSQLExtraction(daoEXTRACTION_018, False, pstrListeMessages, pvarArgument(0))
            
            Select Case modDAO.lngDernièreErreurDAO()
'               Les données sont verrouillées par une transaction
                Case daoENREGISTREMENT_VERROUILLÉ
                    If Not pctlÀDésactiver Is Nothing Then modCommun.DéfinirVerrouillageContrôle pctlÀDésactiver, True
'               Le jeu contient un ou plusieurs enregistrements
                Case daoAUCUNE
                    With prstJeuEnregistrements
                        While Not .EOF
                            modCommun.ImplémenterListe pobjCible, ![DEMA_NUM]
                            .MoveNext
                        Wend
                    End With
            End Select
            
'       Numéros de dossiers
        Case "Dossiers"
        
'           Lire les numéros de dossiers
            Set prstJeuEnregistrements = modDAO.rstSQLExtraction(daoEXTRACTION_017, False, pstrListeMessages)
            
            Select Case modDAO.lngDernièreErreurDAO()
'               Les données sont verrouillées par une transaction
                Case daoENREGISTREMENT_VERROUILLÉ
                    If Not pctlÀDésactiver Is Nothing Then modCommun.DéfinirVerrouillageContrôle pctlÀDésactiver, True
'               Le jeu contient un ou plusieurs enregistrements
                Case daoAUCUNE
                    With prstJeuEnregistrements
                        While Not .EOF
                            modCommun.ImplémenterListe pobjCible, ![DOSS_NUM]
                            .MoveNext
                        Wend
                    End With
            End Select
            
'       Origines ethniques
        Case "Origines ethniques"
        
            modCommun.ImplémenterListe pobjCible, "", pvarNuméro:=adpAUCUNE_SÉLECTION
            
'           Lire les origines ethniques
            Set prstJeuEnregistrements = modDAO.rstSQLExtraction(daoEXTRACTION_021, False, pstrListeMessages)
            
            Select Case modDAO.lngDernièreErreurDAO()
'               Les données sont verrouillées par une transaction
                Case daoENREGISTREMENT_VERROUILLÉ
                    If Not pctlÀDésactiver Is Nothing Then modCommun.DéfinirVerrouillageContrôle pctlÀDésactiver, True
'               Le jeu contient un ou plusieurs enregistrements
                Case daoAUCUNE
                    With prstJeuEnregistrements
                        While Not .EOF
                            modCommun.ImplémenterListe pobjCible, ![ORET_LIBELLE], pvarNuméro:=![ORET_CODE]
                            .MoveNext
                        Wend
                    End With
            End Select
            
'       Provenances d'un agrément
        Case "Provenances"
        
            modCommun.ImplémenterListe pobjCible, "", pvarNuméro:=adpINDÉTERMINÉE
            
            For Each pclsCodeLibellé In gclsProvenances
                With pclsCodeLibellé
                    modCommun.ImplémenterListe pobjCible, .strLibellé, pvarNuméro:=.lngCode
                End With
            Next pclsCodeLibellé
    
'       Types de recours contentieux
        Case "Recours contentieux"
        
            For Each pclsCodeLibellé In gclsRecoursContentieux.clsRecoursContentieux
                With pclsCodeLibellé
                    modCommun.ImplémenterListe pobjCible, .strLibellé, pvarNuméro:=.lngCode
                End With
            Next pclsCodeLibellé
    
'       Sexes
        Case "Sexes"
        
            Select Case pvarArgument(0)
'               La valeur <Indéterminée> doit être affichée
                Case True
                    If UBound(pvarArgument) = 0 Then
                        ReDim pstrValeurs(2)
                        pstrValeurs(0) = "Indéterminé": pstrValeurs(1) = "Masculin": pstrValeurs(2) = "Féminin"
                    Else
                        pstrValeurs = Split(pvarArgument(1), ",")
                    End If
                    modCommun.ImplémenterListe pobjCible, pstrValeurs(0), pvarNuméro:=AdoptionsSexes.adpINDÉTERMINÉ
                    modCommun.ImplémenterListe pobjCible, pstrValeurs(1), pvarNuméro:=adpMASCULIN
                    modCommun.ImplémenterListe pobjCible, pstrValeurs(2), pvarNuméro:=adpFÉMININ
'               La valeur <Indéterminée> ne doit pas être affichée
                Case False
                    If UBound(pvarArgument) = 0 Then
                        ReDim pstrValeurs(1)
                        pstrValeurs(0) = "Masculin": pstrValeurs(1) = "Féminin"
                    Else
                        pstrValeurs = Split(pvarArgument(1), ",")
                    End If
                    modCommun.ImplémenterListe pobjCible, pstrValeurs(0), pvarNuméro:=adpMASCULIN
                    modCommun.ImplémenterListe pobjCible, pstrValeurs(1), pvarNuméro:=adpFÉMININ
            End Select

    End Select
    
    Exit Sub

Err_ImplémenterListe:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Initialiser l'application

Public Sub Initialiser()
Const pkstrNomProcédure As String = "Initialiser"
Const pkstrInitialisationImpossible As String = "L'application ne peut être initialisée."
On Error GoTo Err_Initialiser

    Screen.MousePointer = vbHourglass

'   Ouvrir la connexion ODBC
    modDAO.OuvrirConnexion

'   Implémenter les collections
    modDAO.ImplémenterCollections pkstrInitialisationImpossible & gstrSÉPARATEUR_2 & CStr(vbCritical), "Paramètres de l'échéancier"
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then End
    modDAO.ImplémenterCollections pkstrInitialisationImpossible & gstrSÉPARATEUR_2 & CStr(vbCritical), "Recours contentieux"
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then End
    ImplémenterCollections "Délibérations", "Cadres d'intervention"
    
    Screen.MousePointer = vbDefault

    Exit Sub

Err_Initialiser:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Supprimer les classes initialisées par l'application

Public Sub SupprimerClasses()
Const pkstrNomProcédure As String = "SupprimerClasses"
On Error GoTo Err_SupprimerClasses

'   Supprimer les classes liées aux données
    modDAO.SupprimerClasses
    
'   Supprimer les collections
    Set gclsDélibérations = Nothing: Set gclsCadresInterventions = Nothing

    Exit Sub

Err_SupprimerClasses:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Retourner le nom d'un intervenant
'   pintIntervenantNuméro:      numéro de l'intervenant
'   pstrMessageSiVerrouillage:  message à afficher si les données sont verrouillées par une transaction

Public Function txtNomIntervenant(pintIntervenantNuméro As Integer, pstrMessageSiVerrouillage As String) As String
Const pkstrNomProcédure As String = "txtNomIntervenant"
Dim prstIntervenant As Recordset
On Error GoTo Err_txtNomIntervenant

'   Lire l'identité de l'intervenant
    Set prstIntervenant = modDAO.rstSQLExtraction(daoEXTRACTION_026, True, pstrMessageSiVerrouillage, pintIntervenantNuméro)
    
    Select Case modDAO.lngDernièreErreurDAO()
'       Les données sont verrouillées par une transaction
        Case daoENREGISTREMENT_VERROUILLÉ
            txtNomIntervenant = "<Non lu(e)>"
'       Le jeu ne contient pas d'enregistrement
        Case daoAUCUN_ENREGISTREMENT
            Err.Raise mesMESSAGE_012
            txtNomIntervenant = "<Inconnu(e)>"
'       Le jeu contient un ou plusieurs enregistrements
        Case daoAUCUNE
            txtNomIntervenant = Trim(prstIntervenant![TRSO_NOM])
    End Select
    
    Exit Function
    
Err_txtNomIntervenant:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:     pintTraitement = mesINTERROMPRE
        Case mesMESSAGE_012:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, CStr(pintIntervenantNuméro))
        Case Else:              pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Function


'Retourner le numéro de dossier d'une demande
'   plngNuméroDemande:          numéro de la demande
'   pctlÀDésactiver:            contrôle à désactiver dans le cas où les données sont verrouillées par une transaction
'   pstrMessageSiVerrouillage:  message à afficher si les données sont verrouillées par une transaction

Public Function txtNuméroDossier(plngNuméroDemande As Long, pctlÀDésactiver As Control, Optional pstrMessageSiVerrouillage As String = gstrCONSULTATION_LIMITÉE) As String
Const pkstrNomProcédure As String = "txtNuméroDossier"
Dim prstDemande As Recordset
On Error GoTo Err_txtNuméroDossier

'   Lire le numéro de dossier
    Set prstDemande = modDAO.rstSQLExtraction(daoEXTRACTION_029, False, pstrMessageSiVerrouillage & gstrSÉPARATEUR_1 & gstrDEMANDE_SUPPRIMÉE, plngNuméroDemande)
            
    Select Case modDAO.lngDernièreErreurDAO()
'       Les données sont verrouillées par une transaction
        Case daoENREGISTREMENT_VERROUILLÉ
            If Not pctlÀDésactiver Is Nothing Then modCommun.DéfinirVerrouillageContrôle pctlÀDésactiver, True
'       Le jeu contient un ou plusieurs enregistrements
        Case daoAUCUNE
            txtNuméroDossier = prstDemande![DOSS_NUM]
    End Select
    
    Exit Function
    
Err_txtNuméroDossier:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Function


'Vérifier la cohérence entre les informations présentes sur le recours administratif ou sur la surveillance de placement et celles de l'historique des intervenants
'   plngNuméroDemande:              numéro de la demande
'   pintNuméroRapport:              numéro de rapport indexé à la date de la demande
'   plngCadreIntervention:          cadre d'intervention (recours administratif ou surveillance de placement)
'   pvarPsychologueNuméro:          numéro du psychologue
'   pvarPsychologueDemandé:         date de demande du rapport du psychologue
'   pvarPsychologuePrévu:           date prévue pour la réception du rapport du psychologue
'   pvarPsychologueReçu:            date de réception du rapport du psychologue
'   pvarAssistanteSocialeNuméro:    numéro de l'assistante sociale
'   pvarAssistanteSocialeDemandé:   date de demande du rapport de l'assistante sociale
'   pvarAssistanteSocialePrévu:     date prévue pour la réception du rapport de l'assistante sociale
'   pvarAssistanteSocialeReçu:      date de réception du rapport de l'assistante sociale

Public Sub VérifierCohérenceAvecHistoriqueIntervenants(plngNuméroDemande As Long, pintNuméroRapport As Integer, plngCadreIntervention As AdoptionsCadresInterventions, pvarPsychologueNuméro As Variant, pvarPsychologueDemandé As Variant, pvarPsychologuePrévu As Variant, pvarPsychologueReçu As Variant, pvarAssistanteSocialeNuméro As Variant, pvarAssistanteSocialeDemandé As Variant, pvarAssistanteSocialePrévu As Variant, pvarAssistanteSocialeReçu As Variant)
Const pkstrNomProcédure As String = "VérifierCohérenceAvecHistoriqueIntervenants"
Dim pvarSignetPsychologue As Variant, pvarSignetAssistanteSociale As Variant, pvarSignetIntervenant As Variant, pvarIntervenantNuméro As Variant, pvarIntervenantDemandé As Variant, pvarIntervenantPrévu As Variant, pvarIntervenantReçu As Variant
Dim prstHistoriqueIntervenantsSurCadre As Recordset
Dim pintItération As Integer
Dim pbooIncohérenceAvecHistorique As Boolean, pbooIncohérencePsychologue As Boolean, pbooIncohérenceAssistanteSociale As Boolean, pbooIncohérenceIntervenant As Boolean, pbooAnnulerMiseÀJour As Boolean
On Error GoTo Err_VérifierCohérenceAvecHistoriqueIntervenants

    pvarSignetPsychologue = Null: pvarSignetAssistanteSociale = Null
    
'   Lire l'historique des intervenants liés au cadre d'intervention
    Set prstHistoriqueIntervenantsSurCadre = modDAO.rstSQLExtraction(daoEXTRACTION_027, True, "La vérification de correspondance entre les intervenants et leur historique est annulée." & gstrSÉPARATEUR_2 & CStr(vbInformation), plngNuméroDemande, plngCadreIntervention, pintNuméroRapport)
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub
    
    With prstHistoriqueIntervenantsSurCadre
    
'       Pour chaque intervenant
        While Not .EOF
        
            pbooIncohérenceIntervenant = False
            
            Select Case ![u01_booHistoIntervenantPsychologue]
'               S'il s'agit d'un psychologue
                Case True:  pvarSignetIntervenant = pvarSignetPsychologue:          pvarIntervenantNuméro = pvarPsychologueNuméro:          pvarIntervenantDemandé = pvarPsychologueDemandé:        pvarIntervenantPrévu = pvarPsychologuePrévu:        pvarIntervenantReçu = pvarPsychologueReçu
'               S'il s'agit d'une assistante sociale
                Case False: pvarSignetIntervenant = pvarSignetAssistanteSociale:    pvarIntervenantNuméro = pvarAssistanteSocialeNuméro:    pvarIntervenantDemandé = pvarAssistanteSocialeDemandé:  pvarIntervenantPrévu = pvarAssistanteSocialePrévu:  pvarIntervenantReçu = pvarAssistanteSocialeReçu
            End Select
                    
'           Si l'historique doit être recherché pour ce type d'intervenant (psychologue ou assistante sociale) et qu'il n'a pas encore été trouvé
            If Not IsNull(pvarIntervenantNuméro) And IsNull(pvarSignetIntervenant) Then
'               S'il s'agit de l'historique recherché
                If ![u01_dtmHistoIntervenantDemande] = pvarIntervenantDemandé Then
                    pvarSignetIntervenant = .Bookmark
                    pbooIncohérenceIntervenant = Not (modCommun.varValeurFormatéeSiNull(![u01_dtmHistoIntervenantPrevu], vbDate) = modCommun.varValeurFormatéeSiNull(pvarIntervenantPrévu, vbDate) And modCommun.varValeurFormatéeSiNull(![u01_dtmHistoIntervenantRecu], vbDate) = modCommun.varValeurFormatéeSiNull(pvarIntervenantReçu, vbDate) And ![TRSO_NUM] = pvarIntervenantNuméro)
                End If
            End If
                    
            Select Case ![u01_booHistoIntervenantPsychologue]
'               S'il s'agit d'un psychologue
                Case True:  pvarSignetPsychologue = pvarSignetIntervenant:          If pbooIncohérenceIntervenant Then pbooIncohérencePsychologue = True
'               S'il s'agit d'une assistante sociale
                Case False: pvarSignetAssistanteSociale = pvarSignetIntervenant:    If pbooIncohérenceIntervenant Then pbooIncohérenceAssistanteSociale = True
            End Select
                    
            .MoveNext
            
        Wend
        
    End With

    pbooIncohérencePsychologue = pbooIncohérencePsychologue Or (Not IsNull(pvarPsychologueNuméro) And IsNull(pvarSignetPsychologue))
    pbooIncohérenceAssistanteSociale = pbooIncohérenceAssistanteSociale Or (Not IsNull(pvarAssistanteSocialeNuméro) And IsNull(pvarSignetAssistanteSociale))
    pbooIncohérenceAvecHistorique = pbooIncohérencePsychologue Or pbooIncohérenceAssistanteSociale
    
'   Si les informations sont incohérentes avec l'historique
    If pbooIncohérenceAvecHistorique Then

'       Si l'utilisateur décide la mise à jour de l'historique
        If modMessages.lngRéponseUtilisateur(mesQUESTION_009) = vbYes Then
        
'           Ouvrir une transaction
            modDAO.GérerTransaction daoOUVERTURE
            
            With prstHistoriqueIntervenantsSurCadre
            
'               Pour chaque type d'intervenant (psychologue et assistante sociale)
                For pintItération = 1 To 2
                
                    Select Case pintItération
'                       Psychologue
                        Case 1: pbooIncohérenceIntervenant = pbooIncohérencePsychologue:        pvarSignetIntervenant = pvarSignetPsychologue:          pvarIntervenantNuméro = pvarPsychologueNuméro:          pvarIntervenantDemandé = pvarPsychologueDemandé:        pvarIntervenantPrévu = pvarPsychologuePrévu:        pvarIntervenantReçu = pvarPsychologueReçu
'                       Assistante sociale
                        Case 2: pbooIncohérenceIntervenant = pbooIncohérenceAssistanteSociale:  pvarSignetIntervenant = pvarSignetAssistanteSociale:    pvarIntervenantNuméro = pvarAssistanteSocialeNuméro:    pvarIntervenantDemandé = pvarAssistanteSocialeDemandé:  pvarIntervenantPrévu = pvarAssistanteSocialePrévu:  pvarIntervenantReçu = pvarAssistanteSocialeReçu
                    End Select
                    
'                   Si l'incohérence porte sur le type d'intervenant
                    If pbooIncohérenceIntervenant Then
                    
'                       Si l'historique correspondant n'existe pas
                        If IsNull(pvarSignetIntervenant) Then
                        
                            .AddNew
        
                            ![DEMA_NUM] = plngNuméroDemande
                            ![u01_intHistoIntervenantCadre] = plngCadreIntervention
                            ![u01_booHistoIntervenantPsychologue] = (pintItération = 1)
                            ![u01_intHistoIntervenantNumero] = pintNuméroRapport
        
'                       Sinon
                        Else
                        
                            .Bookmark = pvarSignetIntervenant
                            
                            .Edit
                                    
                        End If
                    
                        ![u01_dtmHistoIntervenantDemande] = pvarIntervenantDemandé
                        ![u01_dtmHistoIntervenantPrevu] = pvarIntervenantPrévu
                        ![u01_dtmHistoIntervenantRecu] = pvarIntervenantReçu
                        ![TRSO_NUM] = pvarIntervenantNuméro
                        
'                       Valider l'opération
                        modDAO.ValiderOpérationSurRecordset prstHistoriqueIntervenantsSurCadre, pstrListeMessages:=gstrSÉPARATEUR_1 & gstrSÉPARATEUR_1 & gstrSÉPARATEUR_1 & gstrDOUBLON_SANS_MESSAGE, pvarRéponseÀInviteÉcrasement:=vbNo
                        Select Case modDAO.lngDernièreErreurDAO()
#If Not TablesEnLocalOuAttachées Then
'                           La création a été refusée du fait de l'ouverture, par un autre utilisateur, de la base ou de l'objet en mode lecture seule
                            Case daoERREUR_01809:               pbooAnnulerMiseÀJour = True
#End If
'                           L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                            Case daoENREGISTREMENT_VERROUILLÉ:  pbooAnnulerMiseÀJour = True
'                           L'opération a été refusée pour éviter la création d'un doublon
                            Case daoDOUBLON:                    pbooAnnulerMiseÀJour = True
'                           L'opération a été refusée en raison de la modification par un autre utilisateur des données originales
                            Case dbRecordModified:              pbooAnnulerMiseÀJour = True
'                           L'opération a été refusée en raison de la suppression par un autre utilisateur des données originales
                            Case dbRecordDeleted:               pbooAnnulerMiseÀJour = True
'                           L'opération a généré un conflit inattendu
                            Case Is <> daoAUCUNE
                                modDAO.GérerTransaction daoANNULATION
                                Err.Raise mesCONFLIT_ACCÈS
                        End Select
                        If pbooAnnulerMiseÀJour Then
                            Err.Raise mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION
                            modDAO.GérerTransaction daoANNULATION: Exit Sub
                        End If
    
                    End If
            
                Next pintItération
                
            End With
                
'           Valider la transaction
            modDAO.GérerTransaction daoVALIDATION
        
        End If
            
    End If
    
    Exit Sub
    
Err_VérifierCohérenceAvecHistoriqueIntervenants:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                                     pintTraitement = mesINTERROMPRE
        Case mesCONFLIT_ACCÈS:                                  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, modDAO.lngDernièreErreurDAO())
        Case mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, "La mise à jour,e", vbInformation)
        Case Else:                                              pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


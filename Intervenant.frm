VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIntervenant 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intervenant - "
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6825
   Begin VB.Frame fraAdoptions 
      Caption         =   "Rapport"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   6615
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   18
         ToolTipText     =   "Date de demande du rapport"
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   16318467
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   3
         Left            =   1215
         TabIndex        =   20
         ToolTipText     =   "Date de réception du rapport"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   16318467
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   4
         Left            =   4860
         TabIndex        =   22
         ToolTipText     =   "Date prévue pour la réception du rapport"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   16318467
         CurrentDate     =   2958465
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Prévu le"
         Height          =   195
         Index           =   7
         Left            =   4155
         TabIndex        =   21
         Top             =   765
         Width           =   585
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Reçu le"
         Height          =   195
         Index           =   6
         Left            =   540
         TabIndex        =   19
         Top             =   765
         Width           =   555
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Demandé le"
         Height          =   195
         Index           =   5
         Left            =   225
         TabIndex        =   17
         Top             =   405
         Width           =   855
      End
   End
   Begin VB.Frame fraAdoptions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   6615
      Begin VB.CommandButton cmdAdoptions 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6060
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Sélectionner l'intervenant"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.TextBox txtAdoptions 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   12
         ToolTipText     =   "Intervenant"
         Top             =   360
         Width           =   4815
      End
      Begin VB.ComboBox cboAdoptions 
         Height          =   315
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Cadre d'intervention"
         Top             =   720
         Width           =   2715
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cadre"
         Height          =   195
         Index           =   4
         Left            =   675
         TabIndex        =   14
         Top             =   765
         Width           =   420
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Identité"
         Height          =   195
         Index           =   3
         Left            =   555
         TabIndex        =   11
         Top             =   405
         Width           =   525
      End
   End
   Begin VB.Frame fraAdoptions 
      Caption         =   "Demande"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6615
      Begin VB.TextBox txtAdoptions 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   7
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   16318467
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   1
         Left            =   4860
         TabIndex        =   9
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   16318467
         CurrentDate     =   2958465
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Référence pour le calcul de l'échéancier"
         Height          =   195
         Index           =   2
         Left            =   3570
         TabIndex        =   8
         Top             =   390
         Width           =   2865
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Déposée le"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   6
         Top             =   765
         Width           =   810
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Numéro"
         Height          =   195
         Index           =   0
         Left            =   525
         TabIndex        =   4
         Top             =   390
         Width           =   555
      End
   End
   Begin VB.Frame fraAdoptions 
      Height          =   945
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton cmdAdoptions 
         Default         =   -1  'True
         Height          =   615
         Index           =   1
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Valider"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton cmdAdoptions 
         Cancel          =   -1  'True
         Height          =   615
         Index           =   0
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Annuler"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar stbAdoptions 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   5355
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "08/10/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15:30"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5812
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmIntervenant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Validation(plngTraitement As CommunTraitements, pbooPsychologue As Boolean, pdtmRapportDemandé As Date, pintRapportNuméro As Integer)

Private WithEvents mfrm_TS As frm_TS
Attribute mfrm_TS.VB_VarHelpID = -1

Private Enum IndexComboBox      'Index des contrôles ComboBox
    cboCADRE = 0
End Enum
Private Enum IndexCommandButton 'Index des contrôles CommandButton
    cmdANNULER = 0
    cmdVALIDER
    cmdINTERVENANT
End Enum
Private Enum IndexDTPicker      'Index des contrôles DTPicker
    dtpDÉPOSÉE = 0
    dtpRÉFÉRENCE
    dtpDEMANDÉ
    dtpREÇU
    dtpPRÉVU
End Enum
Private Enum IndexFrame         'Index des contrôles Frame
    fraACTIONS = 0
    fraDEMANDE
    fraINTERVENANT
    fraRAPPORT
End Enum
Private Enum IndexTextBox       'Index des contrôles TextBox
    txtNUMÉRO = 0
    txtINTERVENANT
End Enum

Private mvarIntervenantNuméro As Variant, mvarSignetTraité As Variant
Private mrstIntervenants As Recordset
Private mdtmRapportDemandé As Date, mdtmRéceptionDossier As Date
Private mlngTraitement As CommunTraitements
Private mlngNuméroDemande As Long
Private mintRapportNuméro As Integer
Private mbooPsychologue As Boolean

Private Sub cmdAdoptions_Click(Index As Integer)
Const pkstrNomProcédure As String = "cmdAdoptions_Click"
Dim pctlFocus As Control
Dim pvarDécision As Variant
Dim pclsCadreIntervention As clsCodeLibellé
Dim pstrMessages(3) As String, pstrValidationAnnulée As String
Dim plngCadre As AdoptionsCadresInterventions
Dim plngParamètreÉchéancier As DAOParamètresÉchéancier
Dim pbooSaisieCorrespondÀHistorique As Boolean
On Error GoTo Err_cmdAdoptions_Click

    Select Case Index

'       Annuler
        Case cmdANNULER
        
            Unload Me

'       Valider
        Case cmdVALIDER

'           Forcer si nécessaire l'actualisation de la valeur d'un contrôle DTPicker
            modCommun.GérerActualisationDTPicker Me, cmdAdoptions(cmdVALIDER)

'           Rafraîchir le jeu des intervenants
            modDAO.ActualiserRecordset mrstIntervenants, True, gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_2 & CStr(vbInformation)
            If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub

'           S'il s'agit d'une suppression
            If mlngTraitement = comSUPPRESSION Then

'               Demander confirmation de la suppression de l'intervenant
                Select Case modMessages.lngRéponseUtilisateur(mesSUPPRIMER, IIf(mbooPsychologue, "le psychologue", "l'assistante sociale"))
'                   L'utilisateur confirme la suppression
                    Case vbYes
'                   L'utilisateur renonce à la suppression
                    Case vbNo
                        cmdAdoptions(cmdANNULER).SetFocus
                        Exit Sub
                End Select

'               Si l'intervenant à supprimer existe bien
                If booPositionnementIntervenantEffectué() Then

                    With mrstIntervenants

'                       Ouvrir une transaction
                        modDAO.GérerTransaction daoOUVERTURE

'                       Valider l'opération
                        modDAO.ValiderOpérationSurRecordset mrstIntervenants, pbooSupprimer:=True, pstrListeMessages:=gstrMESSAGE_SUR_VERROUILLAGE_EN_SUPPRESSION
                        Select Case modDAO.lngDernièreErreurDAO()
'                           L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                            Case daoENREGISTREMENT_VERROUILLÉ
                                modDAO.GérerTransaction daoANNULATION: Exit Sub
'                           L'opération a généré un conflit inattendu
                            Case Is <> daoAUCUNE
                                modDAO.GérerTransaction daoANNULATION
                                Err.Raise mesCONFLIT_ACCÈS
                        End Select

                    End With

                End If

'           Sinon
            Else

                pstrMessages(0) = IIf(mlngTraitement = comCRÉATION, "Création", "Mise à jour")
                
'               Refuser la validation si l'intervenant n'est pas identifié
                If IsNull(mvarIntervenantNuméro) Then
                    Set pctlFocus = cmdAdoptions(cmdINTERVENANT): pstrMessages(1) = IIf(mbooPsychologue, "du psychologue", "de l'assistante sociale")
                    Err.Raise mesSÉLECTION_MANQUANTE
                    Exit Sub
                End If

'               Refuser la validation si le cadre d'intervention n'est pas défini
                If cboAdoptions(cboCADRE).ItemData(cboAdoptions(cboCADRE).ListIndex) = adpAUCUNE_SÉLECTION Then
                    Set pctlFocus = cboAdoptions(cboCADRE): pstrMessages(1) = "du cadre d'intervention"
                    Err.Raise mesSÉLECTION_MANQUANTE
                    Exit Sub
                End If
    
                plngParamètreÉchéancier = IIf(mbooPsychologue, daoRÉCEPTION_DOSSIER__RÉPONSE_PSYCHOLOGUE, daoRÉCEPTION_DOSSIER__RÉPONSE_TRAVAILLEUR_SOCIAL)
                
'               Si la date prévue pour la réception du rapport n'est pas définie
                If IsNull(dtpAdoptions(dtpPRÉVU)) Then
'                   La calculer
'                       Note: La date de référence pour le calcul de l'échéancier aura nécessairement été définie, éventuellement par un positionnement préalable à partir de la date de réception du dossier
                    dtpAdoptions(dtpPRÉVU) = DateAdd("m", modDAO.intParamètreÉchéancier(plngParamètreÉchéancier), dtpAdoptions(dtpRÉFÉRENCE))
                End If

'               Refuser la validation si les informations relatives à l'intervenant sont incohérentes
                If modAdoptions.booSaisieIncohérenteIntervenant(dtpAdoptions(dtpRÉFÉRENCE), mdtmRéceptionDossier, mvarIntervenantNuméro, cmdAdoptions(cmdINTERVENANT), dtpAdoptions(dtpDEMANDÉ), dtpAdoptions(dtpREÇU), IIf(mbooPsychologue, "du psychologue", "de l'assistante sociale"), "de référence" & vbCrLf & "pour le calcul de l'échéancier", "de réception du dossier", pdtpPrévu:=dtpAdoptions(dtpPRÉVU)) Then Exit Sub
                
                plngCadre = cboAdoptions(cboCADRE).ItemData(cboAdoptions(cboCADRE).ListIndex)

                With mrstIntervenants

'                   Pour chaque historique
                    While Not .EOF
                    
                        Select Case mlngTraitement
'                           Création de l'intervenant
                            Case comCRÉATION:       pbooSaisieCorrespondÀHistorique = False
'                           Mise à jour de l'intervenant
                            Case comMODIFICATION:   pbooSaisieCorrespondÀHistorique = (mdtmRapportDemandé = ![u01_dtmHistoIntervenantDemande] And mintRapportNuméro = ![u01_intHistoIntervenantNumero])
                        End Select
                        
'                       Si les informations saisies ne correspondent pas à celles de l'historique et si le cadre d'intervention sélectionné correspond à une demande initiale ou à une seconde évaluation
                        If Not pbooSaisieCorrespondÀHistorique And (plngCadre = AdoptionsCadresInterventions.adpDEMANDE_INITIALE Or plngCadre = adpSECONDE_ÉVALUATION) Then
'                           Refuser la validation si ce cadre d'intervention existe déjà
                            If plngCadre = ![u01_intHistoIntervenantCadre] Then
                                Set pclsCadreIntervention = gclsCadresInterventions(CStr(plngCadre))
                                Set pctlFocus = cboAdoptions(cboCADRE): pstrMessages(1) = "L'historique lié à " & pclsCadreIntervention.strUn & LCase(pclsCadreIntervention.strLibellé)
                                Err.Raise mesEXISTE_DÉJÀ
                                Exit Sub
                            End If
                        End If
                        
'                       Si les informations saisies ne correspondent pas à celles de l'historique et si le cadre d'intervention sélectionné correspond à une demande initiale
                        If Not pbooSaisieCorrespondÀHistorique And plngCadre = AdoptionsCadresInterventions.adpDEMANDE_INITIALE Then
'                           Refuser la validation si la date de demande saisie est postérieure ou égale à la date de demande du rapport de l'historique
                            If Not dtpAdoptions(dtpDEMANDÉ) < ![u01_dtmHistoIntervenantDemande] Then
                                Set pctlFocus = dtpAdoptions(dtpDEMANDÉ): pstrMessages(1) = "de demande saisie": pstrMessages(2) = "de demande d'un historique (" & Format(![u01_dtmHistoIntervenantDemande], gstrFORMAT_DATE_FRANÇAIS) & ")": pstrMessages(3) = IIf(dtpAdoptions(dtpDEMANDÉ) = ![u01_dtmHistoIntervenantDemande], "égale", "postérieure")
                                Err.Raise mesDATE_POSTÉRIEURE
                                Exit Sub
                            End If
                        End If
                        
'                       Si les informations saisies ne correspondent pas à celles de l'historique et si le cadre d'intervention de l'historique correspond à une demande initiale
                        If Not pbooSaisieCorrespondÀHistorique And ![u01_intHistoIntervenantCadre] = AdoptionsCadresInterventions.adpDEMANDE_INITIALE Then
'                           Refuser la validation si la date de demande du rapport de l'historique est postérieure ou égale à la date de demande saisie
                            If Not ![u01_dtmHistoIntervenantDemande] < dtpAdoptions(dtpDEMANDÉ) Then
                                Set pctlFocus = dtpAdoptions(dtpDEMANDÉ): pstrMessages(1) = "de demande de rapport de l'historique correspondant à la demande initiale (" & Format(![u01_dtmHistoIntervenantDemande], gstrFORMAT_DATE_FRANÇAIS) & ")" & vbCrLf: pstrMessages(2) = "de demande saisie": pstrMessages(3) = IIf(![u01_dtmHistoIntervenantDemande] = dtpAdoptions(dtpDEMANDÉ), "égale", "postérieure")
                                Err.Raise mesDATE_POSTÉRIEURE
                                Exit Sub
                            End If
                        End If
                        
                        .MoveNext
                    
                    Wend
                            
                    Select Case mlngTraitement

'                       Création de l'intervenant
                        Case comCRÉATION

'                           Ouvrir une transaction
                            modDAO.GérerTransaction daoOUVERTURE

                            .AddNew

                            ![DEMA_NUM] = mlngNuméroDemande
                            ![u01_booHistoIntervenantPsychologue] = mbooPsychologue
                            ![u01_intHistoIntervenantNumero] = mintRapportNuméro

'                       Mise à jour de l'intervenant
                        Case comMODIFICATION

'                           Si le recours à mettre à jour n'existe plus, sortir
                            If Not booPositionnementIntervenantEffectué() Then Exit Sub

'                           Ouvrir une transaction
                            modDAO.GérerTransaction daoOUVERTURE

                            .Edit
                            
                    End Select

                    ![u01_intHistoIntervenantCadre] = plngCadre
                    ![u01_dtmHistoIntervenantDemande] = dtpAdoptions(dtpDEMANDÉ)
                    ![u01_dtmHistoIntervenantPrevu] = dtpAdoptions(dtpPRÉVU)
                    ![u01_dtmHistoIntervenantRecu] = dtpAdoptions(dtpREÇU)
                    ![TRSO_NUM] = mvarIntervenantNuméro

'                   Valider l'opération
                    pstrValidationAnnulée = IIf(mlngTraitement = comCRÉATION, gstrMESSAGE_SUR_VERROUILLAGE_EN_CRÉATION, gstrMESSAGE_SUR_VERROUILLAGE_EN_MODIFICATION)
                    modDAO.ValiderOpérationSurRecordset mrstIntervenants, pstrListeMessages:=pstrValidationAnnulée & gstrSÉPARATEUR_1 & gstrMESSAGE_SUR_SUPPRESSION & gstrSÉPARATEUR_1 & pstrValidationAnnulée & gstrSÉPARATEUR_1 & "Création refusée: Une demande de rapport existe déjà pour cette date."
                    Select Case modDAO.lngDernièreErreurDAO()
#If Not TablesEnLocalOuAttachées Then
'                       La création a été refusée du fait de l'ouverture, par un autre utilisateur, de la base ou de l'objet en mode lecture seule
                        Case daoERREUR_01809
                            modDAO.GérerTransaction daoANNULATION
                            modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
                            Exit Sub
#End If
'                       L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                        Case daoENREGISTREMENT_VERROUILLÉ
                            modDAO.GérerTransaction daoANNULATION: Exit Sub
'                       L'opération a été refusée pour éviter la création d'un doublon
                        Case daoDOUBLON
                            modDAO.GérerTransaction daoANNULATION
'                           Publier l'invitation à actualiser
                            RaiseEvent Validation(mlngTraitement, mbooPsychologue, dtpAdoptions(dtpDEMANDÉ), daoENREGISTREMENT_SUPPRIMÉ)
                            Exit Sub
'                       L'opération a été refusée en raison de la modification par un autre utilisateur des données originales
                        Case dbRecordModified
                            modDAO.GérerTransaction daoANNULATION
'                           Publier l'invitation à actualiser
                            RaiseEvent Validation(mlngTraitement, mbooPsychologue, dtpAdoptions(dtpDEMANDÉ), daoENREGISTREMENT_SUPPRIMÉ)
                            modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
                            Exit Sub
'                       L'opération a été refusée en raison de la suppression par un autre utilisateur des données originales
                        Case dbRecordDeleted
                            modDAO.GérerTransaction daoANNULATION
'                           Publier l'invitation à actualiser
                            RaiseEvent Validation(mlngTraitement, mbooPsychologue, dtpAdoptions(dtpDEMANDÉ), daoENREGISTREMENT_SUPPRIMÉ)
                            modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
                            Exit Sub
'                       L'opération a généré un conflit inattendu
                        Case Is <> daoAUCUNE
                            modDAO.GérerTransaction daoANNULATION
                            Err.Raise mesCONFLIT_ACCÈS
                    End Select

                End With

            End If

'           Actualiser les données relatives aux intervenants
            modAdoptions.ActualiserDonnéesIntervenants mlngNuméroDemande, gstrVALIDATION_ANNULÉE
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub

'           Valider la transaction
            modDAO.GérerTransaction daoVALIDATION

'           Publier la validation
            RaiseEvent Validation(mlngTraitement, mbooPsychologue, dtpAdoptions(dtpDEMANDÉ), mintRapportNuméro)

            Unload Me
        
'       Sélectionner l'intervenant
        Case cmdINTERVENANT
        
'           Ouvrir la fenêtre permettant le traitement de l'intervenant
            With mfrm_TS
                .Initialiser False, True, True
                .Show vbModal
            End With
    
    End Select
        
    Exit Sub
    
Err_cmdAdoptions_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:             pintTraitement = mesINTERROMPRE
        Case mesCONFLIT_ACCÈS:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, modDAO.lngDernièreErreurDAO())
        Case mesSÉLECTION_MANQUANTE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, pstrMessages(1), pstrMessages(0) & " refusée")
        Case mesDATE_POSTÉRIEURE:       pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, pstrMessages(1), pstrMessages(2), pstrMessages(0) & " refusée", pstrMessages(3))
        Case mesEXISTE_DÉJÀ:            pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, pstrMessages(0), pstrMessages(1))
        Case Else:                      pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub

Private Sub Form_Unload(Cancel As Integer)
Const pkstrNomProcédure As String = "Form_Unload"
On Error GoTo Err_Form_Unload

    Set mrstIntervenants = Nothing
    Set mfrm_TS = Nothing
    
    Exit Sub
    
Err_Form_Unload:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        

End Sub

Private Sub mfrm_TS_Validation(pvarIntervenantNuméro As Variant, pstrIntervenantNom As String)
Const pkstrNomProcédure As String = "mfrm_TS_Validation"
On Error GoTo Err_mfrm_TS_Validation

    mvarIntervenantNuméro = pvarIntervenantNuméro
    txtAdoptions(txtINTERVENANT) = pstrIntervenantNom

    Exit Sub
    
Err_mfrm_TS_Validation:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Retourner un booléen précisant si un signet a pu être posé sur l'intervenant
'   pbooImplémenter:    booléen précisant si les contrôles doivent être implémentés à partir de l'intervenant

Private Function booPositionnementIntervenantEffectué(Optional pbooImplémenter As Boolean = False) As Boolean
Const pkstrNomProcédure As String = "booPositionnementIntervenantEffectué"
On Error GoTo Err_booPositionnementIntervenantEffectué

    With mrstIntervenants

        If Not .RecordCount = 0 Then .MoveFirst
        
'       Pour chaque intervenant
        Do While Not .EOF

'           S'il s'agit de l'intervenant à mettre à jour ou à supprimer
            If ![u01_booHistoIntervenantPsychologue] = mbooPsychologue And ![u01_dtmHistoIntervenantDemande] = mdtmRapportDemandé And ![u01_intHistoIntervenantNumero] = mintRapportNuméro Then

                mvarSignetTraité = .Bookmark
                booPositionnementIntervenantEffectué = True

'               Si les contrôles doivent être implémentés
                If pbooImplémenter Then

                    cboAdoptions(cboCADRE).ListIndex = modCommun.varRecherche(cboAdoptions(cboCADRE), False, ![u01_intHistoIntervenantCadre])
                    dtpAdoptions(dtpDEMANDÉ) = ![u01_dtmHistoIntervenantDemande]
                    If Not IsNull(![u01_dtmHistoIntervenantRecu]) Then dtpAdoptions(dtpREÇU) = ![u01_dtmHistoIntervenantRecu]
                    If Not IsNull(![u01_dtmHistoIntervenantPrevu]) Then dtpAdoptions(dtpPRÉVU) = ![u01_dtmHistoIntervenantPrevu]
                    
'                   Afficher l'identité de l'intervenant
                    mvarIntervenantNuméro = ![TRSO_NUM]
                    modAdoptions.AfficherIdentitéIntervenant CInt(mvarIntervenantNuméro), txtAdoptions(txtINTERVENANT), gstrCONSULTATION_LIMITÉE
    
                End If

                Exit Do

            End If

            .MoveNext

        Loop

    End With

'   Si l'intervenant à mettre à jour ou à supprimer n'existe plus
    If Not booPositionnementIntervenantEffectué Then

'       Publier l'invitation à actualiser
        RaiseEvent Validation(mlngTraitement, mbooPsychologue, mdtmRapportDemandé, daoENREGISTREMENT_SUPPRIMÉ)

        Err.Raise mesENREGISTREMENT_SUPPRIMÉ
        modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True

    End If

    Exit Function
    
Err_booPositionnementIntervenantEffectué:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                 pintTraitement = mesINTERROMPRE
        Case mesENREGISTREMENT_SUPPRIMÉ:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, IIf(mlngTraitement = comMODIFICATION, "Mise à jour", "Suppression"), "L'intervenant")
        Case Else:                          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Function


'Initialiser la fenêtre
'   plngTraitement:     type de traitement
'   plngNuméroDemande:  numéro de demande
'   pbooPsychologue:    booléen précisant s'il s'agit d'un psychologue
'   pvarRapportDemandé: date de demande du rapport              [obligatoire s'il ne s'agit pas d'une création]
'   pvarRapportNuméro:  numéro du rapport indexé à la demande   [obligatoire s'il ne s'agit pas d'une création]

Public Sub Initialiser(plngTraitement As CommunTraitements, plngNuméroDemande As Long, pbooPsychologue As Boolean, Optional pvarRapportDemandé As Variant, Optional pvarRapportNuméro As Variant)
Const pkstrNomProcédure As String = "Initialiser"
Dim prstDemande As Recordset
On Error GoTo Err_Initialiser
    
    mlngTraitement = plngTraitement: mlngNuméroDemande = plngNuméroDemande: mbooPsychologue = pbooPsychologue

'   Centrer la feuille
    modAdoptions.CentrerFeuille Me
        
'   Implémenter les images
    cmdAdoptions(cmdANNULER).Picture = Menu.picImage(adp32x32, adpANNULER)
    cmdAdoptions(cmdVALIDER).Picture = Menu.picImage(adp32x32, adpVALIDER)
    cmdAdoptions(cmdINTERVENANT).Picture = Menu.picImage(adp16x16, adpINTERVENANTS)
        
'   Désactiver certains contrôles
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtNUMÉRO), True
    modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpDÉPOSÉE), True
    modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpRÉFÉRENCE), True
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtINTERVENANT), True
    
'   Implémenter le titre
    Me.Caption = Me.Caption & Choose(mlngTraitement, "Création", "", "Modification", "Suppression")
    
'   Implémenter certains contrôles
    txtAdoptions(txtNUMÉRO) = CStr(mlngNuméroDemande)
    fraAdoptions(fraINTERVENANT).Caption = IIf(mbooPsychologue, "Psychologue", "Assistante sociale")
    
'   Implémenter le cadre d'intervention
    modAdoptions.ImplémenterListe cboAdoptions(cboCADRE), "Cadres d'interventions", "", Nothing, False
    
'   Positionner les dates de rapport à la date du jour
    modCommun.PositionnerDateJour dtpAdoptions(dtpDEMANDÉ), dtpAdoptions(dtpREÇU), dtpAdoptions(dtpPRÉVU)
    dtpAdoptions(dtpDEMANDÉ) = "": dtpAdoptions(dtpREÇU) = "": dtpAdoptions(dtpPRÉVU) = ""
    
    Set mfrm_TS = New frm_TS
    mvarIntervenantNuméro = Null
    
'   Lire la demande
    Set prstDemande = modDAO.rstSQLExtraction(daoDEMANDE_DATES, True, gstrCONSULTATION_LIMITÉE & gstrSÉPARATEUR_1 & gstrDEMANDE_SUPPRIMÉE, mlngNuméroDemande)
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
        modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
        Exit Sub
    End If
    With prstDemande
        dtpAdoptions(dtpDÉPOSÉE) = ![DEMA_DATE]
        mdtmRéceptionDossier = ![DEMA_DREDOS]
        dtpAdoptions(dtpRÉFÉRENCE) = IIf(IsNull(![DEMA_DREF]), "", ![DEMA_DREF])    'Note: La date peut être nulle dans le cas de la suppression d'un intervenant créé préalablement à l'obligation faite par le logiciel de définir une date de référence pour le traitement des intervenants
    End With
    
'   Vérifier l'existence d'intervenants
    Set mrstIntervenants = modDAO.rstSQLExtraction(daoPSYCHOLOGUES__OU__ASSISTANTES_SOCIALES, True, gstrCONSULTATION_LIMITÉE, dbOpenDynaset, mlngNuméroDemande, mbooPsychologue)
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
        modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
        Exit Sub
    End If

'   S'il s'agit d'une création
    If mlngTraitement = comCRÉATION Then
        mintRapportNuméro = 1
        cboAdoptions(cboCADRE).ListIndex = modCommun.varRecherche(cboAdoptions(cboCADRE), False, adpAUCUNE_SÉLECTION)
'   Sinon
    Else
        mdtmRapportDemandé = pvarRapportDemandé: mintRapportNuméro = pvarRapportNuméro
'       Si l'intervenant à mettre à jour ou à supprimer n'existe plus, sortir
        If Not booPositionnementIntervenantEffectué(True) Then Exit Sub
    End If

'   S'il s'agit d'une suppression
    If mlngTraitement = comSUPPRESSION Then
'       Désactiver la saisie
        modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdINTERVENANT), True
        modCommun.DéfinirVerrouillageContrôle cboAdoptions(cboCADRE), True, True
        modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpDEMANDÉ), True, True
        modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpREÇU), True, True
        modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpPRÉVU), True, True
    End If
            
    Exit Sub
    
Err_Initialiser:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

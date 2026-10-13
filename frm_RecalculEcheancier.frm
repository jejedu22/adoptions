VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_RecalculEcheancier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recalcul de l'échéancier"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fra_Identification 
      Caption         =   "Identification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   8415
      Begin VB.TextBox txtAdoptions 
         Enabled         =   0   'False
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
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   7
         ToolTipText     =   "Demandeur"
         Top             =   360
         Width           =   4245
      End
      Begin VB.ComboBox Cmb_NumDemande 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Numéro de la demande"
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox Cmb_NumDossier 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Numéro du dossier"
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   11
         ToolTipText     =   "Date de réception du dossier"
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
         Format          =   16449539
         CurrentDate     =   2958465
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Demandeur"
         Height          =   195
         Index           =   0
         Left            =   3015
         TabIndex        =   6
         Top             =   405
         Width           =   825
      End
      Begin VB.Label Lbl_DateDemande 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dossier reçu le"
         Height          =   195
         Left            =   2790
         TabIndex        =   10
         Top             =   765
         Width           =   1050
      End
      Begin VB.Label Lbl_NumDemande 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Demande n°"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   765
         Width           =   885
      End
      Begin VB.Label lbl_NumDossier 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dossier n°"
         Height          =   195
         Left            =   285
         TabIndex        =   4
         Top             =   405
         Width           =   720
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   945
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Annuler"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   615
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Valider"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Fra_Pièces 
      Caption         =   "Dates de références pour le calcul de l'échéancier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   8415
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   2
         Left            =   6720
         TabIndex        =   17
         ToolTipText     =   "Nouvelle date de référence"
         Top             =   810
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   16449539
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   1
         Left            =   6720
         TabIndex        =   15
         ToolTipText     =   "Date de référence actuelle"
         Top             =   360
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
         Format          =   16449539
         CurrentDate     =   2958465
      End
      Begin VB.Label Lbl_ReferenceActuelle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Actuelle"
         Height          =   195
         Left            =   5970
         TabIndex        =   14
         Top             =   405
         Width           =   630
      End
      Begin VB.Label Lbl_Notes 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Notes"
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Lbl_DateRecalcul 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nouvelle"
         Height          =   195
         Left            =   5970
         TabIndex        =   16
         Top             =   855
         Width           =   630
      End
   End
   Begin MSComctlLib.StatusBar stb_RecalculEcheancier 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   4245
      Width           =   8700
      _ExtentX        =   15346
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
            TextSave        =   "21/07/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10:45"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9119
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_RecalculEcheancier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum IndexDTPicker      'Index des contrôles DTPicker
    dtpRÉCEPTION_DOSSIER = 0
    dtpRÉFÉRENCE_ACTUELLE
    dtpRÉFÉRENCE_NOUVELLE
End Enum
Private Enum IndexTextBox       'Index des contrôles TextBox
    txtDEMANDEUR = 0
End Enum

Private Const mstrRECALCUL_REFUSÉ As String = "Le recalcul est impossible."

Private QuitterCroix As Boolean
Private mbooInitialisationAvecDemandeDéfinie As Boolean

Private Sub Cmb_NumDemande_Click()
Const pkstrNomProcédure As String = "Cmb_NumDemande_Click"
Dim prstDemande As Recordset
On Error GoTo Err_Cmb_NumDemande_Click

'   Lire la demande
    Set prstDemande = modDAO.rstSQLExtraction(daoDEMANDE_DATES, True, mstrRECALCUL_REFUSÉ & gstrSÉPARATEUR_1 & gstrDEMANDE_SUPPRIMÉE, CLng(Cmb_NumDemande))
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
        modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
        Exit Sub
    End If
    
'   Positionner les dates
    With prstDemande
        dtpAdoptions(dtpRÉCEPTION_DOSSIER).Value = modCommun.varValeurFormatéeSiNull(![DEMA_DREDOS], vbDate)
        dtpAdoptions(dtpRÉFÉRENCE_ACTUELLE).Value = IIf(IsNull(![DEMA_DREF]), dtpAdoptions(dtpRÉCEPTION_DOSSIER).Value, ![DEMA_DREF])
        If dtpAdoptions(dtpRÉCEPTION_DOSSIER).Value = #12/31/9999# Then dtpAdoptions(dtpRÉCEPTION_DOSSIER) = ""
        If dtpAdoptions(dtpRÉFÉRENCE_ACTUELLE).Value = #12/31/9999# Then dtpAdoptions(dtpRÉFÉRENCE_ACTUELLE) = ""
        dtpAdoptions(dtpRÉFÉRENCE_NOUVELLE).Value = dtpAdoptions(dtpRÉFÉRENCE_ACTUELLE).Value: dtpAdoptions(dtpRÉFÉRENCE_NOUVELLE) = ""
    End With
    
'   Désactiver les identifiants
    modCommun.DéfinirVerrouillageContrôle Cmb_NumDossier, True, True
    modCommun.DéfinirVerrouillageContrôle Cmb_NumDemande, True, True
            
    Exit Sub
    
Err_Cmb_NumDemande_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub Cmb_NumDossier_Click()
Const pkstrNomProcédure As String = "Cmb_NumDossier_Click"
On Error GoTo Err_Cmb_NumDossier_Click

'   Afficher l'identité du demandeur
    modAdoptions.AfficherIdentitéDemandeur Cmb_NumDossier, txtAdoptions(txtDEMANDEUR), mstrRECALCUL_REFUSÉ, cmd_Valider
        
'   Si la demande n'est pas définie à l'initialisation de la fenêtre
    If Not mbooInitialisationAvecDemandeDéfinie Then
'       Lire les numéros des demandes liées au dossier
        modAdoptions.ImplémenterListe Cmb_NumDemande, "Demandes", mstrRECALCUL_REFUSÉ, cmd_Valider, Cmb_NumDossier
'       S'il n'en existe qu'une, la sélectionner
        If Cmb_NumDemande.ListCount = 1 Then Cmb_NumDemande.ListIndex = 0
    End If
    
    If Me.Visible And Cmb_NumDemande.Enabled Then Cmb_NumDemande.SetFocus
    
    Exit Sub
    
Err_Cmb_NumDossier_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_Quitter_Click()
    Unload Me
End Sub

Private Sub cmd_Valider_Click()
Const pkstrNomProcédure As String = "cmd_Valider_Click"
Const pkstrRecalculAnnulé As String = "Le recalcul est annulé.", pkstrAucuneDate As String = "<Aucune date>"
Dim prstDécision As Recordset, prstDatesPrévuesRéceptionRapports As Recordset, prstRecoursAdministratif As Recordset, prstIntervenants As Recordset
Dim pfldColonne As Field
Dim pctlFocus As Control
Dim pdtmNouvelleDateRéférence As Date
Dim pstrMessages(1) As String, pstrDateOriginale As String, pstrDateRecalculée As String, pstrDateNature As String
Dim plngIntervenantsTraités As Long, plngIntervenant As AdoptionsIntervenants
Dim pbooAnnulerTransaction As Boolean, pbooRecalculerRecoursAdministratif As Boolean
On Error GoTo Err_cmd_Valider_Click

'   Forcer si nécessaire l'actualisation de la valeur d'un contrôle DTPicker
    modCommun.GérerActualisationDTPicker Me, cmd_Valider

    Set pctlFocus = IIf(Cmb_NumDemande.Enabled, Cmb_NumDemande, cmd_Quitter)
    
'   Refuser la validation si le numéro de la demande n'est pas sélectionné
    If Cmb_NumDemande = "" Then
        pstrMessages(0) = "du numéro de la demande": Err.Raise mesSÉLECTION_MANQUANTE
        Exit Sub
    End If
    
'   Refuser la validation si la date de réception n'est pas définie
    If IsNull(dtpAdoptions(dtpRÉCEPTION_DOSSIER)) Then
        pstrMessages(0) = "le recalcul de l'échéancier": Err.Raise mesDATE_RÉCEPTION_ABSENTE
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
'   Lire la dernière/prochaine commission liée à la demande
    Set prstDécision = modDAO.rstSQLExtraction(daoEXTRACTION_005, False, pkstrRecalculAnnulé, CLng(Cmb_NumDemande))
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Screen.MousePointer = vbDefault: Exit Sub
    With prstDécision
'       Si elle existe
        If Not .RecordCount = 0 Then
'           Si la demande a fait l'objet d'une décision, ne recalculer que les dates se rapportant au recours administratif (s'il existe)
            If Not ![u02_intHistoCommissionDecision] = AdoptionsDélibérations.adpINDÉTERMINÉ Then pbooRecalculerRecoursAdministratif = True
        End If
    End With
    
'   Lire les dates prévues pour la réception des rapports des intervenants
    Set prstDatesPrévuesRéceptionRapports = modDAO.rstSQLExtraction(daoEXTRACTION_006, False, pkstrRecalculAnnulé & gstrSÉPARATEUR_1 & gstrDEMANDE_SUPPRIMÉE, CLng(Cmb_NumDemande))
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then Screen.MousePointer = vbDefault: Exit Sub
    With prstDatesPrévuesRéceptionRapports
'       Si elles sont définies
        If Not (IsNull(![DEMA_DPRRP]) And IsNull(![DEMA_DPRRA])) Then
'           Si le dossier est sorti ou la demande suspendue, ne recalculer que les dates se rapportant au recours administratif (s'il existe)
            If Not (IsNull(![DEMA_DSORT]) And IsNull(![DEMA_DSUSP])) Then pbooRecalculerRecoursAdministratif = True
        End If
    End With
    
'   Déterminer la nouvelle date de référence
    If IsNull(dtpAdoptions(dtpRÉFÉRENCE_NOUVELLE)) Then pdtmNouvelleDateRéférence = dtpAdoptions(dtpRÉFÉRENCE_ACTUELLE).Value Else pdtmNouvelleDateRéférence = dtpAdoptions(dtpRÉFÉRENCE_NOUVELLE).Value
            
'   Ouvrir une transaction
    modDAO.GérerTransaction daoOUVERTURE
        
'   Si le recalcul ne doit porter que sur les dates se rapportant au recours administratif (s'il existe)
    If pbooRecalculerRecoursAdministratif Then
    
        With prstDatesPrévuesRéceptionRapports
        
            .Edit
            
'           Mettre à jour la date de référence
            ![DEMA_DREF] = pdtmNouvelleDateRéférence
            
'           Valider la mise à jour
            modDAO.ValiderOpérationSurRecordset prstDatesPrévuesRéceptionRapports
            Select Case modDAO.lngDernièreErreurDAO()
'               L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                Case daoENREGISTREMENT_VERROUILLÉ:  pbooAnnulerTransaction = True
'               L'opération a été refusée en raison de la suppression par un autre utilisateur des données originales
                Case dbRecordDeleted:               pbooAnnulerTransaction = True
'               L'opération a été refusée en raison de la modification par un autre utilisateur des données originales
                Case dbRecordModified:              pbooAnnulerTransaction = True
'               L'opération a généré un conflit inattendu
                Case Is <> daoAUCUNE:               Err.Raise mesCONFLIT_ACCÈS
            End Select
'           Si la transaction doit être annulée
            If pbooAnnulerTransaction Then
'               L'annuler
                modDAO.GérerTransaction daoANNULATION
                Screen.MousePointer = vbDefault: Exit Sub
            End If
                            
        End With
        
        With prstDécision
        
'           S'il existe une dernière/prochaine commission liée à la demande
            If Not .RecordCount = 0 Then

'               Vérifier l'existence d'un recours administratif pour cette commission
                Set prstRecoursAdministratif = modDAO.rstSQLExtraction(daoEXTRACTION_008, False, pkstrRecalculAnnulé, ![DEMA_NUM], ![u02_dtmHistoCommissionReunion])
                If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
                    modDAO.GérerTransaction daoANNULATION
                    Screen.MousePointer = vbDefault: Exit Sub
                End If
            
'               S'il existe
                If Not modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT Then
                
                    With prstRecoursAdministratif
                        
                        .Edit
                        
'                       Recalculer la date prévue pour la décision
                        ![READ_DPREDEC] = DateAdd("m", modDAO.intParamètreÉchéancier(daoRÉCEPTION_DOSSIER__DÉCISION), ![READ_DDREC])
                        Set pfldColonne = ![READ_DPREDEC]: pstrDateNature = ": (recours administratif) date prévue pour la décision"
                        pstrDateOriginale = Format(pfldColonne.OriginalValue, gstrFORMAT_DATE_FRANÇAIS): pstrDateRecalculée = Format(pfldColonne.Value, gstrFORMAT_DATE_FRANÇAIS): pstrMessages(0) = pstrMessages(0) & vbCrLf & "- " & IIf(pstrDateOriginale = "", pkstrAucuneDate, pstrDateOriginale) & " -> " & IIf(pstrDateRecalculée = "", pkstrAucuneDate, pstrDateRecalculée) & pstrDateNature
'                       Recalculer les dates prévues pour la réception des rapports (si elles existent)
                        If Not IsNull(![READ_DPRRECAS]) Then
                            ![READ_DPRRECAS] = DateAdd("m", modDAO.intParamètreÉchéancier(daoRECOURS_ADMINISTRATIF__RÉPONSE_TRAVAILLEUR_SOCIAL), ![READ_DDEMAS])
                            Set pfldColonne = ![READ_DPRRECAS]: pstrDateNature = ": (recours administratif) date prévue pour la réception du rapport de l'assistante sociale"
                            pstrDateOriginale = Format(pfldColonne.OriginalValue, gstrFORMAT_DATE_FRANÇAIS): pstrDateRecalculée = Format(pfldColonne.Value, gstrFORMAT_DATE_FRANÇAIS): pstrMessages(0) = pstrMessages(0) & vbCrLf & "- " & IIf(pstrDateOriginale = "", pkstrAucuneDate, pstrDateOriginale) & " -> " & IIf(pstrDateRecalculée = "", pkstrAucuneDate, pstrDateRecalculée) & pstrDateNature
                        End If
                        If Not IsNull(![READ_DPRRECPSY]) Then
                            ![READ_DPRRECPSY] = DateAdd("m", modDAO.intParamètreÉchéancier(daoRECOURS_ADMINISTRATIF__RÉPONSE_PSYCHOLOGUE), ![READ_DDEMPSY])
                            Set pfldColonne = ![READ_DPRRECPSY]: pstrDateNature = ": (recours administratif) date prévue pour la réception du rapport du psychologue"
                            pstrDateOriginale = Format(pfldColonne.OriginalValue, gstrFORMAT_DATE_FRANÇAIS): pstrDateRecalculée = Format(pfldColonne.Value, gstrFORMAT_DATE_FRANÇAIS): pstrMessages(0) = pstrMessages(0) & vbCrLf & "- " & IIf(pstrDateOriginale = "", pkstrAucuneDate, pstrDateOriginale) & " -> " & IIf(pstrDateRecalculée = "", pkstrAucuneDate, pstrDateRecalculée) & pstrDateNature
                        End If
                    
'                       Valider la mise à jour
                        modDAO.ValiderOpérationSurRecordset prstRecoursAdministratif
                        Select Case modDAO.lngDernièreErreurDAO()
'                           L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                            Case daoENREGISTREMENT_VERROUILLÉ:  pbooAnnulerTransaction = True
'                           L'opération a été refusée en raison de la suppression par un autre utilisateur des données originales
                            Case dbRecordDeleted:               pbooAnnulerTransaction = True
'                           L'opération a été refusée en raison de la modification par un autre utilisateur des données originales
                            Case dbRecordModified:              pbooAnnulerTransaction = True
'                           L'opération a généré un conflit inattendu
                            Case Is <> daoAUCUNE:               Err.Raise mesCONFLIT_ACCÈS
                        End Select
'                       Si la transaction doit être annulée
                        If pbooAnnulerTransaction Then
'                           L'annuler
                            modDAO.GérerTransaction daoANNULATION
                            Screen.MousePointer = vbDefault: Exit Sub
                        End If
                                
'                       Lire les intervenants
                        Set prstIntervenants = modDAO.rstSQLExtraction(daoEXTRACTION_007, False, pkstrRecalculAnnulé, ![DEMA_NUM], ![READ_DDEMPSY], ![TRSO_NUMPSY], ![READ_DDEMAS], ![TRSO_NUMAS])
                        If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
                            modDAO.GérerTransaction daoANNULATION
                            Screen.MousePointer = vbDefault: Exit Sub
                        End If
            
'                       Pour chaque intervenant
                        While Not prstIntervenants.EOF
            
                            prstIntervenants.Edit
        
'                           Recalculer la date prévue pour la réception du rapport
                            prstIntervenants![u01_dtmHistoIntervenantPrevu] = IIf(prstIntervenants![u01_booHistoIntervenantPsychologue], ![READ_DPRRECPSY], ![READ_DPRRECAS])
                            Set pfldColonne = prstIntervenants![u01_dtmHistoIntervenantPrevu]: pstrDateNature = ": (historique intervenants) date prévue pour la réception du rapport " & IIf(prstIntervenants![u01_booHistoIntervenantPsychologue], "du psychologue", "de l'assistante sociale")
                            If Not IsNull(pfldColonne) Then pstrDateOriginale = Format(pfldColonne.OriginalValue, gstrFORMAT_DATE_FRANÇAIS): pstrDateRecalculée = Format(pfldColonne.Value, gstrFORMAT_DATE_FRANÇAIS): pstrMessages(0) = pstrMessages(0) & vbCrLf & "- " & IIf(pstrDateOriginale = "", pkstrAucuneDate, pstrDateOriginale) & " -> " & IIf(pstrDateRecalculée = "", pkstrAucuneDate, pstrDateRecalculée) & pstrDateNature
        
'                           Valider la mise à jour
                            modDAO.ValiderOpérationSurRecordset prstIntervenants
                            Select Case modDAO.lngDernièreErreurDAO()
'                               L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                                Case daoENREGISTREMENT_VERROUILLÉ:  pbooAnnulerTransaction = True
'                               L'opération a été refusée en raison de la suppression par un autre utilisateur des données originales
                                Case dbRecordDeleted:               pbooAnnulerTransaction = True
'                               L'opération a été refusée en raison de la modification par un autre utilisateur des données originales
                                Case dbRecordModified:              pbooAnnulerTransaction = True
'                               L'opération a généré un conflit inattendu
                                Case Is <> daoAUCUNE:               Err.Raise mesCONFLIT_ACCÈS
                            End Select
'                           Si la transaction doit être annulée
                            If pbooAnnulerTransaction Then
'                               L'annuler
                                modDAO.GérerTransaction daoANNULATION
                                Screen.MousePointer = vbDefault: Exit Sub
                            End If
            
                            prstIntervenants.MoveNext
            
                        Wend
    
                    End With
                    
                End If
                
            End If
        
        End With
    
'   Sinon
    Else
    
        With prstDatesPrévuesRéceptionRapports
        
            .Edit
            
'           Mettre à jour la date de référence
            ![DEMA_DREF] = pdtmNouvelleDateRéférence
'           Recalculer les dates prévues pour la réception des rapports (si elles existent)
            If Not IsNull(![DEMA_DPRRP]) Then
                ![DEMA_DPRRP] = DateAdd("m", modDAO.intParamètreÉchéancier(daoRÉCEPTION_DOSSIER__RÉPONSE_PSYCHOLOGUE), pdtmNouvelleDateRéférence)
                Set pfldColonne = ![DEMA_DPRRP]: pstrDateNature = ": date prévue pour la réception du rapport du psychologue"
                pstrDateOriginale = Format(pfldColonne.OriginalValue, gstrFORMAT_DATE_FRANÇAIS): pstrDateRecalculée = Format(pfldColonne.Value, gstrFORMAT_DATE_FRANÇAIS): pstrMessages(0) = pstrMessages(0) & vbCrLf & "- " & IIf(pstrDateOriginale = "", pkstrAucuneDate, pstrDateOriginale) & " -> " & IIf(pstrDateRecalculée = "", pkstrAucuneDate, pstrDateRecalculée) & pstrDateNature
            End If
            If Not IsNull(![DEMA_DPRRA]) Then
                ![DEMA_DPRRA] = DateAdd("m", modDAO.intParamètreÉchéancier(daoRÉCEPTION_DOSSIER__RÉPONSE_TRAVAILLEUR_SOCIAL), pdtmNouvelleDateRéférence)
                Set pfldColonne = ![DEMA_DPRRA]: pstrDateNature = ": date prévue pour la réception du rapport de l'assistante sociale"
                pstrDateOriginale = Format(pfldColonne.OriginalValue, gstrFORMAT_DATE_FRANÇAIS): pstrDateRecalculée = Format(pfldColonne.Value, gstrFORMAT_DATE_FRANÇAIS): pstrMessages(0) = pstrMessages(0) & vbCrLf & "- " & IIf(pstrDateOriginale = "", pkstrAucuneDate, pstrDateOriginale) & " -> " & IIf(pstrDateRecalculée = "", pkstrAucuneDate, pstrDateRecalculée) & pstrDateNature
            End If
            
'           Valider la mise à jour
            modDAO.ValiderOpérationSurRecordset prstDatesPrévuesRéceptionRapports
            Select Case modDAO.lngDernièreErreurDAO()
'               L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                Case daoENREGISTREMENT_VERROUILLÉ:  pbooAnnulerTransaction = True
'               L'opération a été refusée en raison de la suppression par un autre utilisateur des données originales
                Case dbRecordDeleted:               pbooAnnulerTransaction = True
'               L'opération a été refusée en raison de la modification par un autre utilisateur des données originales
                Case dbRecordModified:              pbooAnnulerTransaction = True
'               L'opération a généré un conflit inattendu
                Case Is <> daoAUCUNE:               Err.Raise mesCONFLIT_ACCÈS
            End Select
'           Si la transaction doit être annulée
            If pbooAnnulerTransaction Then
'               L'annuler
                modDAO.GérerTransaction daoANNULATION
                Screen.MousePointer = vbDefault: Exit Sub
            End If
                            
'           Lire les intervenants
            Set prstIntervenants = modDAO.rstSQLExtraction(daoEXTRACTION_007, False, pkstrRecalculAnnulé, ![DEMA_NUM], ![DEMA_DDEMRP], ![TRSO_NUMPSY], ![DEMA_DDEMRA], ![TRSO_NUMAS])
            If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
                modDAO.GérerTransaction daoANNULATION
                Screen.MousePointer = vbDefault: Exit Sub
            End If

            plngIntervenantsTraités = 0

'           Pour chaque intervenant
            While Not prstIntervenants.EOF

                plngIntervenant = IIf(prstIntervenants![u01_booHistoIntervenantPsychologue], adpPSYCHOLOGUE, adpTRAVAILLEUR_SOCIAL)

'               Si le type d'intervenant n'a pas été traité
'                   Note: Plusieurs enregistrements mettant en jeu le même le type d'intervenant (psychologue ou assistante sociale) peuvent être retournés dans le cas d'une surveillance de placement, la date de la demande étant chaque fois la même; il ne faut alors traiter que le plus récent
                If Not (plngIntervenantsTraités And plngIntervenant) = plngIntervenant Then

                    prstIntervenants.Edit

'                   Recalculer la date prévue pour la réception du rapport
                    prstIntervenants![u01_dtmHistoIntervenantPrevu] = IIf(prstIntervenants![u01_booHistoIntervenantPsychologue], ![DEMA_DPRRP], ![DEMA_DPRRA])
                    Set pfldColonne = prstIntervenants![u01_dtmHistoIntervenantPrevu]: pstrDateNature = ": (historique intervenants) date prévue pour la réception du rapport " & IIf(prstIntervenants![u01_booHistoIntervenantPsychologue], "du psychologue", "de l'assistante sociale")
                    If Not IsNull(pfldColonne) Then pstrDateOriginale = Format(pfldColonne.OriginalValue, gstrFORMAT_DATE_FRANÇAIS): pstrDateRecalculée = Format(pfldColonne.Value, gstrFORMAT_DATE_FRANÇAIS): pstrMessages(0) = pstrMessages(0) & vbCrLf & "- " & IIf(pstrDateOriginale = "", pkstrAucuneDate, pstrDateOriginale) & " -> " & IIf(pstrDateRecalculée = "", pkstrAucuneDate, pstrDateRecalculée) & pstrDateNature

'                   Valider la mise à jour
                    modDAO.ValiderOpérationSurRecordset prstIntervenants
                    Select Case modDAO.lngDernièreErreurDAO()
'                       L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                        Case daoENREGISTREMENT_VERROUILLÉ:  pbooAnnulerTransaction = True
'                       L'opération a été refusée en raison de la suppression par un autre utilisateur des données originales
                        Case dbRecordDeleted:               pbooAnnulerTransaction = True
'                       L'opération a été refusée en raison de la modification par un autre utilisateur des données originales
                        Case dbRecordModified:              pbooAnnulerTransaction = True
'                       L'opération a généré un conflit inattendu
                        Case Is <> daoAUCUNE:               Err.Raise mesCONFLIT_ACCÈS
                    End Select
'                   Si la transaction doit être annulée
                    If pbooAnnulerTransaction Then
'                       L'annuler
                        modDAO.GérerTransaction daoANNULATION
                        Screen.MousePointer = vbDefault: Exit Sub
                    End If

                    plngIntervenantsTraités = plngIntervenantsTraités Or plngIntervenant

                End If

                prstIntervenants.MoveNext

            Wend
    
        End With
        
        With prstDécision
        
'           S'il existe une dernière/prochaine commission liée à la demande
            If Not .RecordCount = 0 Then
            
                .Edit
                
'               Recalculer la date prévue pour la commission
                ![u02_dtmHistoCommissionPrevu] = DateAdd("m", modDAO.intParamètreÉchéancier(daoRÉCEPTION_DOSSIER__DÉCISION), pdtmNouvelleDateRéférence)
                Set pfldColonne = ![u02_dtmHistoCommissionPrevu]: pstrDateNature = ": date prévue pour la décision"
                pstrDateOriginale = Format(pfldColonne.OriginalValue, gstrFORMAT_DATE_FRANÇAIS): pstrDateRecalculée = Format(pfldColonne.Value, gstrFORMAT_DATE_FRANÇAIS): pstrMessages(0) = pstrMessages(0) & vbCrLf & "- " & IIf(pstrDateOriginale = "", pkstrAucuneDate, pstrDateOriginale) & " -> " & IIf(pstrDateRecalculée = "", pkstrAucuneDate, pstrDateRecalculée) & pstrDateNature
    
'               Valider la mise à jour
                modDAO.ValiderOpérationSurRecordset prstDécision
                Select Case modDAO.lngDernièreErreurDAO()
'                   L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                    Case daoENREGISTREMENT_VERROUILLÉ:  pbooAnnulerTransaction = True
'                   L'opération a été refusée en raison de la suppression par un autre utilisateur des données originales
                    Case dbRecordDeleted:               pbooAnnulerTransaction = True
'                   L'opération a été refusée en raison de la modification par un autre utilisateur des données originales
                    Case dbRecordModified:              pbooAnnulerTransaction = True
'                   L'opération a généré un conflit inattendu
                    Case Is <> daoAUCUNE:               Err.Raise mesCONFLIT_ACCÈS
                End Select
'               Si la transaction doit être annulée
                If pbooAnnulerTransaction Then
'                   L'annuler
                    modDAO.GérerTransaction daoANNULATION
                    Screen.MousePointer = vbDefault: Exit Sub
                End If
    
            End If
        
        End With
    
    End If

    Screen.MousePointer = vbDefault
    
'   Valider la transaction
    modDAO.GérerTransaction daoVALIDATION
        
'   Informer l'utilisateur de la nature des dates recalculées
    If pstrMessages(0) = "" Then pstrMessages(0) = vbCrLf & "- <Aucune>"
    Err.Raise mesMESSAGE_002
    
    Unload Me

    Exit Sub
  
Err_cmd_Valider_Click:
Dim pintTraitement As MessagesActions
    Select Case Err
        Case vbObjectError:             pintTraitement = mesINTERROMPRE
        Case mesCONFLIT_ACCÈS:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, modDAO.lngDernièreErreurDAO())
        Case mesSÉLECTION_MANQUANTE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, pstrMessages(0))
        Case mesDATE_RÉCEPTION_ABSENTE: pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, pstrMessages(0))
        Case mesMESSAGE_002:            pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, pstrMessages(0))
        Case Else:                      pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    Screen.MousePointer = vbDefault
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub

Private Sub dtpAdoptions_GotFocus(Index As Integer)
Const pkstrNomProcédure As String = "dtpAdoptions_GotFocus"
On Error GoTo Err_dtpAdoptions_GotFocus

    Select Case Index
        Case dtpRÉFÉRENCE_NOUVELLE
            stb_RecalculEcheancier.SimpleText = "Entrez la date pour le recalcul de l'échéancier"
    End Select

    Exit Sub
    
Err_dtpAdoptions_GotFocus:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Initialiser la fenêtre
'   pvarNuméroDemande:  numéro de demande [facultatif]

Public Sub Initialiser(Optional pvarNuméroDemande As Variant)
Const pkstrNomProcédure As String = "Initialiser"
Dim prstDemandeur As Recordset
Dim pstrNuméroDossier As String
On Error GoTo Err_Initialiser
    
    mbooInitialisationAvecDemandeDéfinie = Not IsMissing(pvarNuméroDemande)
    
'   Centrer la feuille
    modAdoptions.CentrerFeuille Me
        
'   Implémenter les images
    cmd_Quitter.Picture = Menu.picImage(adp32x32, adpANNULER)
    cmd_Valider.Picture = Menu.picImage(adp32x32, adpVALIDER)
    
'   Désactiver certains contrôles
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtDEMANDEUR), True
    modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpRÉCEPTION_DOSSIER), True
    modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpRÉFÉRENCE_ACTUELLE), True
    
    dtpAdoptions(dtpRÉFÉRENCE_NOUVELLE) = ""

    Lbl_Notes.Caption = "Note:" & vbCrLf & "Si aucune nouvelle date n'est saisie et cochée pour le recalcul, c'est la date de référence actuelle qui sera utilisée."
    
'   Si la demande est définie
    If mbooInitialisationAvecDemandeDéfinie Then
        pstrNuméroDossier = modAdoptions.txtNuméroDossier(CLng(pvarNuméroDemande), cmd_Valider, mstrRECALCUL_REFUSÉ)
        If pstrNuméroDossier = "" Then Exit Sub
        modCommun.ImplémenterListe Cmb_NumDossier, pstrNuméroDossier
        modCommun.ImplémenterListe Cmb_NumDemande, CStr(pvarNuméroDemande)
        Cmb_NumDossier.ListIndex = modCommun.varRecherche(Cmb_NumDossier, True, pstrNuméroDossier)
        Cmb_NumDemande.ListIndex = modCommun.varRecherche(Cmb_NumDemande, True, CStr(pvarNuméroDemande))
'   Sinon
    Else
        modAdoptions.ImplémenterListe Cmb_NumDossier, "Dossiers", mstrRECALCUL_REFUSÉ, cmd_Valider
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


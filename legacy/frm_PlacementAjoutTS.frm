VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_PlacementAjoutTS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Surveillance du placement - "
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8520
   StartUpPosition =   1  'CenterOwner
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
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   8295
      Begin VB.TextBox txtAdoptions 
         Alignment       =   2  'Center
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
         Left            =   1740
         TabIndex        =   18
         ToolTipText     =   "Assistante sociale"
         Top             =   840
         Width           =   2295
      End
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
         Index           =   1
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Sélectionner le psychologue"
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   375
      End
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
         Index           =   0
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Sélectionner l'assistante sociale"
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.TextBox txtAdoptions 
         Alignment       =   2  'Center
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
         Left            =   1740
         TabIndex        =   23
         ToolTipText     =   "Psychologue"
         Top             =   1320
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   2
         Left            =   4695
         TabIndex        =   20
         ToolTipText     =   "Date d'intervention de l'assistante sociale"
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59113475
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   4
         Left            =   4695
         TabIndex        =   25
         ToolTipText     =   "Date d'intervention du psychologue"
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59113475
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   3
         Left            =   6510
         TabIndex        =   21
         ToolTipText     =   "Date de réception du rapport de l'assistante sociale"
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59113475
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   5
         Left            =   6510
         TabIndex        =   26
         ToolTipText     =   "Date de réception du rapport du psychologue"
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59113475
         CurrentDate     =   2958465
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Intervenants"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   1080
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   2  'Center
         Caption         =   "Rapport reçu le"
         Height          =   240
         Index           =   7
         Left            =   6510
         TabIndex        =   16
         Top             =   465
         Width           =   1575
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   2  'Center
         Caption         =   "Intervention le"
         Height          =   255
         Index           =   6
         Left            =   4695
         TabIndex        =   15
         Top             =   465
         Width           =   1575
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FF0000&
         X1              =   6390
         X2              =   6390
         Y1              =   480
         Y2              =   1680
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FF0000&
         X1              =   4575
         X2              =   4575
         Y1              =   480
         Y2              =   1680
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   2  'Center
         Caption         =   "Nom"
         Height          =   255
         Index           =   5
         Left            =   1740
         TabIndex        =   14
         Top             =   465
         Width           =   2295
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF0000&
         X1              =   1620
         X2              =   1620
         Y1              =   480
         Y2              =   1680
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Psychologue"
         Height          =   195
         Index           =   9
         Left            =   585
         TabIndex        =   22
         Top             =   1350
         Width           =   915
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         X1              =   240
         X2              =   8085
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Assistante sociale"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   17
         Top             =   870
         Width           =   1260
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         X1              =   240
         X2              =   8085
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         X1              =   240
         X2              =   8085
         Y1              =   720
         Y2              =   720
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
      ForeColor       =   &H00FF0000&
      Height          =   1305
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   8295
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   9
         ToolTipText     =   "Date de la demande de surveillance"
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59113475
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   1
         Left            =   1455
         TabIndex        =   11
         ToolTipText     =   "Date d'envoi du dossier au Tribunal de Grande Instance"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59113475
         CurrentDate     =   2958465
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Surveillance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   1065
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Demandée le"
         Height          =   195
         Index           =   2
         Left            =   375
         TabIndex        =   8
         Top             =   405
         Width           =   945
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Envoi au TGI le"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   10
         Top             =   765
         Width           =   1110
      End
   End
   Begin VB.Frame fraAdoptions 
      Caption         =   "Placement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   945
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   8295
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
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Numéro"
         Height          =   195
         Index           =   0
         Left            =   765
         TabIndex        =   4
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame fraAdoptions 
      Height          =   945
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   615
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Valider"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Annuler"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar stb_AjoutTS 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   5745
      Width           =   8520
      _ExtentX        =   15028
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
            TextSave        =   "16:28"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6923
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_PlacementAjoutTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Validation(plngTraitement As CommunTraitements, pdtmSurveillance As Date, pintNuméroRapport As Integer)

Private WithEvents mfrm_TS As frm_TS
Attribute mfrm_TS.VB_VarHelpID = -1

Private Enum IndexCommandButton             'Index des contrôles CommandButton
    cmdASSISTANTE_SOCIALE = 0
    cmdPSYCHOLOGUE
End Enum
Private Enum IndexDTPicker                  'Index des contrôles DTPicker
    dtpSURVEILLANCE = 0
    dtpENVOI_TGI
    dtpASSISTANTE_SOCIALE_INTERVENTION
    dtpASSISTANTE_SOCIALE_REÇU
    dtpPSYCHOLOGUE_INTERVENTION
    dtpPSYCHOLOGUE_REÇU
End Enum
Private Enum IndexFrame                     'Index des contrôles Frame
    fraACTIONS = 0
    fraPLACEMENT
    fraSURVEILLANCE
    fraINTERVENANTS
End Enum
Private Enum IndexTextBox                   'Index des contrôles TextBox
    txtPLACEMENT = 0
    txtASSISTANTE_SOCIALE
    txtPSYCHOLOGUE
End Enum

Private mvarAssistanteSocialeNuméro As Variant, mvarPsychologueNuméro As Variant, mvarDateRéceptionPrévue As Variant
Private mvarSurveillance As Variant
Private mrstSurveillances As Recordset
Private mdtmSurveillance As Date, mdtmRéceptionDossier As Date
Private mlngTraitement As CommunTraitements
Private mlngIntervenant As IndexCommandButton
Private mlngNuméroPlacement As Long, mlngNuméroDemande As Long
Private mintNuméroRapport As Integer

Private Sub cmd_Quitter_Click()
Const pkstrNomProcédure As String = "cmd_Quitter_Click"
On Error GoTo Err_cmd_Quitter_Click
        
    Unload Me

    Exit Sub
    
Err_cmd_Quitter_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
End Sub

Private Sub cmd_Valider_Click()
Const pkstrNomProcédure As String = "cmd_Valider_Click"
Dim pvarNuméro As Variant
Dim prstNombreDemandesAutresQueSurveillances As Recordset
Dim pctlFocus As Control
Dim pcmdIntervenant As CommandButton
Dim pdtpReçu As DTPicker, pdtpIntervention As DTPicker
Dim pstrMessages(2) As String, pstrChaîne As String, pstrValidationAnnulée As String
Dim plngDernièreErreurDAO As DAOErreurs
Dim pintItération As Integer, pintNuméroRapport As Integer
On Error GoTo Err_cmd_Valider_Click

'   Forcer si nécessaire l'actualisation de la valeur d'un contrôle DTPicker
    modCommun.GérerActualisationDTPicker Me, cmd_Valider

'   Rafraîchir le jeu des surveillances
    modDAO.ActualiserRecordset mrstSurveillances, True, gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_2 & CStr(vbInformation)
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub

'   S'il s'agit d'une suppression
    If mlngTraitement = comSUPPRESSION Then

'       Demander confirmation de la suppression de la surveillance
        Select Case modMessages.lngRéponseUtilisateur(mesSUPPRIMER, "la surveillance")
'           L'utilisateur confirme la suppression
            Case vbYes
'           L'utilisateur renonce à la suppression
            Case vbNo
                cmd_Quitter.SetFocus
                Exit Sub
        End Select

'       Si la surveillance à supprimer existe bien
        If booPositionnementSurveillanceEffectué() Then

            With mrstSurveillances

'               Ouvrir une transaction
                modDAO.GérerTransaction daoOUVERTURE

'               Pour chaque intervenant
                For pintItération = 1 To 2
'                   Supprimer l'historique
                    modDAO.SQLSupprimer daoDEL_HISTORIQUE_INTERVENANT, True, gstrVALIDATION_ANNULÉE, mlngNuméroDemande, (pintItération = 2), dtpAdoptions(dtpSURVEILLANCE), mintNuméroRapport
                    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                Next pintItération
                    
'               Valider l'opération
                modDAO.ValiderOpérationSurRecordset mrstSurveillances, pbooSupprimer:=True, pstrListeMessages:=gstrMESSAGE_SUR_VERROUILLAGE_EN_SUPPRESSION
                Select Case modDAO.lngDernièreErreurDAO()
'                   L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                    Case daoENREGISTREMENT_VERROUILLÉ
                        modDAO.GérerTransaction daoANNULATION: Exit Sub
'                   L'opération a généré un conflit inattendu
                    Case Is <> daoAUCUNE
                        modDAO.GérerTransaction daoANNULATION
                        Err.Raise mesCONFLIT_ACCÈS
                End Select

            End With

        End If

'   Sinon
    Else

        pstrMessages(0) = IIf(mlngTraitement = comCRÉATION, "Création", "Mise à jour")
                
'       Refuser la validation si la date de la demande de surveillance n'est pas définie
        If IsNull(dtpAdoptions(dtpSURVEILLANCE)) Then
            Set pctlFocus = dtpAdoptions(dtpSURVEILLANCE): pstrMessages(0) = "de la date de la demande de surveillance"
            Err.Raise mesSAISIE_MANQUANTE
            Exit Sub
        End If

'       Refuser la validation si la date d'envoi du dossier au TGI est antérieure à la date de la demande de surveillance
        If modCommun.varValeurFormatéeSiNull(dtpAdoptions(dtpENVOI_TGI), vbDate) < dtpAdoptions(dtpSURVEILLANCE) Then
            Set pctlFocus = dtpAdoptions(dtpENVOI_TGI): pstrMessages(0) = "d'envoi du dossier au Tribunal de Grande Instance": pstrMessages(1) = "de la demande de surveillance"
            Err.Raise mesDATE_ANTÉRIEURE
            Exit Sub
        End If

'       Si la date prévue pour la réception des rapports n'est pas définie
        If IsNull(mvarDateRéceptionPrévue) Then
'           La calculer
            mvarDateRéceptionPrévue = DateAdd("m", modDAO.intParamètreÉchéancier(daoDEUX_SURVEILLANCES_AVANT_JUGEMENT), dtpAdoptions(dtpSURVEILLANCE))
        End If
        
'       Ouvrir une transaction
        modDAO.GérerTransaction daoOUVERTURE
    
'       Pour chaque intervenant
        For pintItération = 1 To 2
    
            Select Case pintItération
'               Assistante sociale
                Case 1: pvarNuméro = mvarAssistanteSocialeNuméro:   Set pcmdIntervenant = cmdAdoptions(cmdASSISTANTE_SOCIALE):  Set pdtpReçu = dtpAdoptions(dtpASSISTANTE_SOCIALE_REÇU):    Set pdtpIntervention = dtpAdoptions(dtpASSISTANTE_SOCIALE_INTERVENTION):    pstrChaîne = "de l'assistante sociale"
'               Psychologue
                Case 2: pvarNuméro = mvarPsychologueNuméro:         Set pcmdIntervenant = cmdAdoptions(cmdPSYCHOLOGUE):         Set pdtpReçu = dtpAdoptions(dtpPSYCHOLOGUE_REÇU):           Set pdtpIntervention = dtpAdoptions(dtpPSYCHOLOGUE_INTERVENTION):           pstrChaîne = "du psychologue"
            End Select
            
'           Refuser la validation si l'intervenant n'est pas identifié
            If IsNull(pvarNuméro) Then
                Set pctlFocus = pcmdIntervenant: pstrMessages(1) = pstrChaîne
                Err.Raise mesSÉLECTION_MANQUANTE
                modDAO.GérerTransaction daoANNULATION: Exit Sub
            End If

'           Refuser la validation si les informations relatives à l'intervenant sont incohérentes
            If modAdoptions.booSaisieIncohérenteIntervenant(dtpAdoptions(dtpSURVEILLANCE), mdtmRéceptionDossier, pvarNuméro, pcmdIntervenant, dtpAdoptions(dtpSURVEILLANCE), pdtpReçu, pstrChaîne, "de la demande de surveillance", "de réception du dossier", pdtpIntervention:=pdtpIntervention) Then
                modDAO.GérerTransaction daoANNULATION: Exit Sub
            End If
            
'           S'il s'agit d'une création
            If mlngTraitement = comCRÉATION Then

'               Créer l'historique
                Do
                
                    If pintItération = 1 Then pintNuméroRapport = pintNuméroRapport + 1
                    
                    modDAO.SQLAjouter daoADD_HISTORIQUE_INTERVENANT, True, gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_1 & gstrDOUBLON_SANS_MESSAGE, mlngNuméroDemande, adpSURVEILLANCE_PLACEMENT, (pintItération = 2), dtpAdoptions(dtpSURVEILLANCE), pintNuméroRapport, mvarDateRéceptionPrévue, pdtpReçu, pvarNuméro
                    plngDernièreErreurDAO = modDAO.lngDernièreErreurDAO()
                    If Not (plngDernièreErreurDAO = daoAUCUNE Or plngDernièreErreurDAO = daoDOUBLON) Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                    
'                   Si une demande de rapport existe déjà pour cette date
                    If plngDernièreErreurDAO = daoDOUBLON Then
                    
'                       Vérifier qu'elle ne se rapporte pas à autre chose qu'à une surveillance de placement
                        Set prstNombreDemandesAutresQueSurveillances = modDAO.rstSQLExtraction(daoEXTRACTION_032, True, gstrVALIDATION_ANNULÉE, mlngNuméroDemande, adpSURVEILLANCE_PLACEMENT, (pintItération = 2), dtpAdoptions(dtpSURVEILLANCE))
                        If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                
'                       Si ce n'est pas le cas
                        If Not prstNombreDemandesAutresQueSurveillances![Nombre] = 0 Then
                            Set pctlFocus = dtpAdoptions(dtpSURVEILLANCE): pstrMessages(0) = "Création": pstrMessages(1) = "Une demande de rapport ne se rapportant pas" & vbCrLf & "à une surveillance": pstrMessages(2) = "pour cette date"
                            Err.Raise mesEXISTE_DÉJÀ
                            modDAO.GérerTransaction daoANNULATION: Exit Sub
                        End If
                    
                    End If
                    
                Loop While plngDernièreErreurDAO = daoDOUBLON
                
                mintNuméroRapport = pintNuméroRapport

'           Sinon
            Else

'               Mettre à jour l'historique
                modDAO.SQLMettreÀJour daoMAJ_HISTORIQUE_INTERVENANT, True, gstrVALIDATION_ANNULÉE, mlngNuméroDemande, (pintItération = 2), dtpAdoptions(dtpSURVEILLANCE), mintNuméroRapport, dtpAdoptions(dtpSURVEILLANCE), mvarDateRéceptionPrévue, pdtpReçu, pvarNuméro
                If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub

            End If
                
        Next pintItération
    
        With mrstSurveillances

            Select Case mlngTraitement

'               Création de la surveillance
                Case comCRÉATION

                    .AddNew

                    ![PLAC_NUM] = mlngNuméroPlacement
                    ![SUPL_DDSUR] = dtpAdoptions(dtpSURVEILLANCE)
                    ![SUPL_intNumero] = mintNuméroRapport

'               Mise à jour de la surveillance
                Case comMODIFICATION

'                   Si la surveillance à mettre à jour n'existe plus, sortir
                    If Not booPositionnementSurveillanceEffectué() Then modDAO.GérerTransaction daoANNULATION: Exit Sub

                    .Edit

            End Select

            ![TRSO_NUM_AssistanteSociale] = mvarAssistanteSocialeNuméro
            ![SUPL_dtmAssistanteSocialeIntervention] = dtpAdoptions(dtpASSISTANTE_SOCIALE_INTERVENTION)
            ![SUPL_dtmAssistanteSocialeRecu] = dtpAdoptions(dtpASSISTANTE_SOCIALE_REÇU)
            ![SUPL_DENVTGI] = dtpAdoptions(dtpENVOI_TGI)
            ![SUPL_dtmAssistanteSocialePrevu] = mvarDateRéceptionPrévue
            ![TRSO_NUM_Psychologue] = mvarPsychologueNuméro
            ![SUPL_dtmPsychologueIntervention] = dtpAdoptions(dtpPSYCHOLOGUE_INTERVENTION)
            ![SUPL_dtmPsychologueRecu] = dtpAdoptions(dtpPSYCHOLOGUE_REÇU)
            ![SUPL_dtmPsychologuePrevu] = mvarDateRéceptionPrévue

'           Valider l'opération
            pstrValidationAnnulée = IIf(mlngTraitement = comCRÉATION, gstrMESSAGE_SUR_VERROUILLAGE_EN_CRÉATION, gstrMESSAGE_SUR_VERROUILLAGE_EN_MODIFICATION)
            modDAO.ValiderOpérationSurRecordset mrstSurveillances, pstrListeMessages:=pstrValidationAnnulée & gstrSÉPARATEUR_1 & gstrMESSAGE_SUR_SUPPRESSION & gstrSÉPARATEUR_1 & pstrValidationAnnulée & gstrSÉPARATEUR_1 & "Création refusée: Une demande de rapport existe déjà pour cette date."
            Select Case modDAO.lngDernièreErreurDAO()
#If Not TablesEnLocalOuAttachées Then
'               La création a été refusée du fait de l'ouverture, par un autre utilisateur, de la base ou de l'objet en mode lecture seule
                Case daoERREUR_01809
                    modDAO.GérerTransaction daoANNULATION
                    Implémenter "<interdiction de valider>", "Vrai"
                    Exit Sub
#End If
'               L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                Case daoENREGISTREMENT_VERROUILLÉ
                    modDAO.GérerTransaction daoANNULATION: Exit Sub
'               L'opération a été refusée pour éviter la création d'un doublon
'                   Note: Le cas ne devrait théoriquement jamais se produire, le doublon ayant été intercepté lors de l'écriture sur la table u01_HistoIntervenants
                Case daoDOUBLON
                    modDAO.GérerTransaction daoANNULATION
'                   Publier l'invitation à actualiser
                    RaiseEvent Validation(mlngTraitement, dtpAdoptions(dtpSURVEILLANCE), daoENREGISTREMENT_SUPPRIMÉ)
                    Exit Sub
'               L'opération a été refusée en raison de la modification par un autre utilisateur des données originales
                Case dbRecordModified
                    modDAO.GérerTransaction daoANNULATION
'                   Publier l'invitation à actualiser
                    RaiseEvent Validation(mlngTraitement, dtpAdoptions(dtpSURVEILLANCE), daoENREGISTREMENT_SUPPRIMÉ)
                    Implémenter "<interdiction de valider>", "Vrai"
                    Exit Sub
'               L'opération a été refusée en raison de la suppression par un autre utilisateur des données originales
                Case dbRecordDeleted
                    modDAO.GérerTransaction daoANNULATION
'                   Publier l'invitation à actualiser
                    RaiseEvent Validation(mlngTraitement, dtpAdoptions(dtpSURVEILLANCE), daoENREGISTREMENT_SUPPRIMÉ)
                    Implémenter "<interdiction de valider>", "Vrai"
                    Exit Sub
'               L'opération a généré un conflit inattendu
                Case Is <> daoAUCUNE
                    modDAO.GérerTransaction daoANNULATION
                    Err.Raise mesCONFLIT_ACCÈS
            End Select

        End With

    End If

'   Actualiser les données relatives aux intervenants
    modAdoptions.ActualiserDonnéesIntervenants mlngNuméroDemande, gstrVALIDATION_ANNULÉE
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub

'   Valider la transaction
    modDAO.GérerTransaction daoVALIDATION

'   Publier la validation
    RaiseEvent Validation(mlngTraitement, dtpAdoptions(dtpSURVEILLANCE), mintNuméroRapport)

    Unload Me
        
    Exit Sub
    
Err_cmd_Valider_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:             pintTraitement = mesINTERROMPRE
        Case mesCONFLIT_ACCÈS:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, modDAO.lngDernièreErreurDAO())
        Case mesSÉLECTION_MANQUANTE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, pstrMessages(1), pstrMessages(0) & " refusée")
        Case mesSAISIE_MANQUANTE:       pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, pstrMessages(0))
        Case mesDATE_ANTÉRIEURE:        pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, pstrMessages(0), pstrMessages(1))
        Case mesEXISTE_DÉJÀ:            pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, pstrMessages(0), pstrMessages(1), pstrMessages(2))
        Case Else:                      pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
  
End Sub

Private Sub cmdAdoptions_Click(Index As Integer)
Const pkstrNomProcédure As String = "cmdAdoptions_Click"
On Error GoTo Err_cmdAdoptions_Click

'   Ouvrir la fenêtre permettant le traitement de l'intervenant
    With mfrm_TS
        mlngIntervenant = Index
        .Initialiser False, False, True
        .Show vbModal
    End With
    
    Exit Sub
    
Err_cmdAdoptions_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub dtpAdoptions_Change(Index As Integer)
Const pkstrNomProcédure As String = "dtpAdoptions_Change"
On Error GoTo Err_dtpAdoptions_Change

    Select Case Index
    
'       Demandée le
        Case dtpSURVEILLANCE
        
'           Activer/Désactiver les contrôles se rapportant à la surveillance
            Implémenter "<activation de la surveillance>"
            
    End Select
    
    Exit Sub
    
Err_dtpAdoptions_Change:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const pkstrNomProcédure As String = "Form_Unload"
On Error GoTo Err_Form_Unload

    Set mrstSurveillances = Nothing
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

    Select Case mlngIntervenant
        Case cmdASSISTANTE_SOCIALE: mvarAssistanteSocialeNuméro = pvarIntervenantNuméro:    txtAdoptions(txtASSISTANTE_SOCIALE) = pstrIntervenantNom
        Case cmdPSYCHOLOGUE:        mvarPsychologueNuméro = pvarIntervenantNuméro:          txtAdoptions(txtPSYCHOLOGUE) = pstrIntervenantNom
    End Select
    
    Exit Sub
    
Err_mfrm_TS_Validation:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Retourner un booléen précisant si un signet a pu être posé sur la surveillance
'   pbooImplémenter:    booléen précisant si les contrôles doivent être implémentés à partir de la surveillance

Private Function booPositionnementSurveillanceEffectué(Optional pbooImplémenter As Boolean = False) As Boolean
Const pkstrNomProcédure As String = "booPositionnementSurveillanceEffectué"
On Error GoTo Err_booPositionnementSurveillanceEffectué

    With mrstSurveillances
    
        If Not .RecordCount = 0 Then .MoveFirst
        
'       Pour chaque surveillance
        Do While Not .EOF
        
'           S'il s'agit de la surveillance à mettre à jour ou à supprimer
            If ![SUPL_DDSUR] = mdtmSurveillance And ![SUPL_intNumero] = mintNuméroRapport Then

                mvarSurveillance = .Bookmark
                booPositionnementSurveillanceEffectué = True

'               Si les contrôles doivent être implémentés
                If pbooImplémenter Then
                
                    dtpAdoptions(dtpSURVEILLANCE) = ![SUPL_DDSUR]
                    dtpAdoptions(dtpENVOI_TGI) = modCommun.varValeurFormatéeSiNull(![SUPL_DENVTGI], vbDate, Null)
                    dtpAdoptions(dtpASSISTANTE_SOCIALE_INTERVENTION) = modCommun.varValeurFormatéeSiNull(![SUPL_dtmAssistanteSocialeIntervention], vbDate, Null)
                    dtpAdoptions(dtpASSISTANTE_SOCIALE_REÇU) = modCommun.varValeurFormatéeSiNull(![SUPL_dtmAssistanteSocialeRecu], vbDate, Null)
                    dtpAdoptions(dtpPSYCHOLOGUE_INTERVENTION) = modCommun.varValeurFormatéeSiNull(![SUPL_dtmPsychologueIntervention], vbDate, Null)
                    dtpAdoptions(dtpPSYCHOLOGUE_REÇU) = modCommun.varValeurFormatéeSiNull(![SUPL_dtmPsychologueRecu], vbDate, Null)
                    
'                   Afficher l'identité des intervenants
                    mvarAssistanteSocialeNuméro = ![TRSO_NUM_AssistanteSociale]:    modAdoptions.AfficherIdentitéIntervenant CInt(mvarAssistanteSocialeNuméro), txtAdoptions(txtASSISTANTE_SOCIALE), gstrCONSULTATION_LIMITÉE
                    mvarPsychologueNuméro = ![TRSO_NUM_Psychologue]:                modAdoptions.AfficherIdentitéIntervenant CInt(mvarPsychologueNuméro), txtAdoptions(txtPSYCHOLOGUE), gstrCONSULTATION_LIMITÉE
                    
                    mvarDateRéceptionPrévue = ![SUPL_dtmAssistanteSocialePrevu]
                    
                End If

                Exit Do

            End If
            
            .MoveNext
            
        Loop
        
    End With

'   Si la surveillance à mettre à jour ou à supprimer n'existe plus
    If Not booPositionnementSurveillanceEffectué Then
    
'       Publier l'invitation à actualiser
        RaiseEvent Validation(mlngTraitement, mdtmSurveillance, CInt(daoENREGISTREMENT_SUPPRIMÉ))
        
        Err.Raise mesENREGISTREMENT_SUPPRIMÉ
        
'       Interdire la validation
        Implémenter "<interdiction de valider>", "Vrai"
            
    End If

    Exit Function
    
Err_booPositionnementSurveillanceEffectué:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                 pintTraitement = mesINTERROMPRE
        Case mesENREGISTREMENT_SUPPRIMÉ:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, IIf(mlngTraitement = comMODIFICATION, "Mise à jour", "Suppression"), "La surveillance")
        Case Else:                          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Function



'Implémenter les objets
'   pstrListeObjets:        liste des objets à implémenter au format '<objet#1>[|<objet#2>[...|<objet#n]]'
'   pstrListeParamètres:    liste des paramètres au format '<paramètre#1 de l'objet#1[¤paramètre#2 de l'objet#1[...¤paramètre#n de l'objet#1]]>[|<paramètre#1 de l'objet#2[¤paramètre#2 de l'objet#2[...¤paramètre#n de l'objet#2]]>[...|<paramètre#1 de l'objet#n[¤paramètre#2 de l'objet#n[...¤paramètre#n de l'objet#n]]>]]'

Private Sub Implémenter(pstrListeObjets As String, Optional pstrListeParamètres As String = "")
Const pkstrNomProcédure As String = "Implémenter"
Dim pstrListesObjets() As String, pstrListesParamètres() As String, pstrParamètres() As String
Dim pintItération As Integer
Dim pbooVerrouillé As Boolean
On Error GoTo Err_Implémenter
    
    pstrListesObjets() = Split(pstrListeObjets, ","): pstrListesParamètres() = Split(pstrListeParamètres, gstrSÉPARATEUR_1)
    
'   Pour chaque objet dont on souhaite l'implémentation
    For pintItération = 0 To UBound(pstrListesObjets)
    
        Select Case pstrListesObjets(pintItération)
        
'           Gestion de l'activation des contrôles se rapportant à la surveillance
            Case "<activation de la surveillance>"
        
                pbooVerrouillé = IsNull(dtpAdoptions(dtpSURVEILLANCE))
                
                modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpENVOI_TGI), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdASSISTANTE_SOCIALE), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpASSISTANTE_SOCIALE_INTERVENTION), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpASSISTANTE_SOCIALE_REÇU), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdPSYCHOLOGUE), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpPSYCHOLOGUE_INTERVENTION), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpPSYCHOLOGUE_REÇU), pbooVerrouillé
                
'           Gestion de l'activation des frames
            Case "<désactivation des frames>"

                pstrParamètres() = Split(pstrListesParamètres(pintItération), gstrSÉPARATEUR_2)
                pbooVerrouillé = CBool(pstrParamètres(0))

                modCommun.DéfinirVerrouillageContrôle fraAdoptions(fraSURVEILLANCE), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle fraAdoptions(fraINTERVENANTS), pbooVerrouillé

'           Interdiction de valider
            Case "<interdiction de valider>"
        
                pstrParamètres() = Split(pstrListesParamètres(pintItération), gstrSÉPARATEUR_2)
                pbooVerrouillé = CBool(pstrParamètres(0))
                
                modCommun.DéfinirVerrouillageContrôle cmd_Valider, pbooVerrouillé
                
                If pbooVerrouillé Then
'                   Désactiver les frames
                    Implémenter "<désactivation des frames>", "Vrai"
                End If

        End Select
        
    Next pintItération
    
    Exit Sub
    
Err_Implémenter:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Initialiser la fenêtre
'   plngTraitement:         type de traitement
'   plngNuméroPlacement:    numéro de placement
'   pvarSurveillance:       date de la demande de surveillance                                  [obligatoire s'il ne s'agit pas d'une création]
'   pvarNuméroRapport:      Numéro de rapport indexé à la date de la demande de surveillance    [obligatoire s'il ne s'agit pas d'une création]

Public Sub Initialiser(plngTraitement As CommunTraitements, plngNuméroPlacement As Long, Optional pvarSurveillance As Variant, Optional pvarNuméroRapport As Variant)
Const pkstrNomProcédure As String = "Initialiser"
Dim prstPlacement As Recordset
Dim pstrMessages(0) As String
On Error GoTo Err_Initialiser
    
    mlngTraitement = plngTraitement: mlngNuméroPlacement = plngNuméroPlacement
    mvarSurveillance = Null

'   Centrer la feuille
    modAdoptions.CentrerFeuille Me
        
'   Implémenter les images
    cmd_Quitter.Picture = Menu.picImage(adp32x32, adpANNULER)
    cmd_Valider.Picture = Menu.picImage(adp32x32, adpVALIDER)
    cmdAdoptions(cmdASSISTANTE_SOCIALE).Picture = Menu.picImage(adp16x16, adpINTERVENANTS)
    cmdAdoptions(cmdPSYCHOLOGUE).Picture = Menu.picImage(adp16x16, adpINTERVENANTS)
        
'   Désactiver certains contrôles
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtPLACEMENT), True
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtASSISTANTE_SOCIALE), True
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtPSYCHOLOGUE), True
    
'   Implémenter le titre
    Me.Caption = Me.Caption & Choose(mlngTraitement, "Création", "", "Modification", "Suppression")
    
'   Positionner les références du placement
    txtAdoptions(txtPLACEMENT) = CStr(plngNuméroPlacement)
    
    Set mfrm_TS = New frm_TS
    mvarAssistanteSocialeNuméro = Null: mvarPsychologueNuméro = Null: mvarDateRéceptionPrévue = Null
    mintNuméroRapport = 0
    
'   Positionner les dates à la date du jour
    modCommun.PositionnerDateJour dtpAdoptions(dtpSURVEILLANCE), dtpAdoptions(dtpENVOI_TGI), dtpAdoptions(dtpASSISTANTE_SOCIALE_INTERVENTION), dtpAdoptions(dtpASSISTANTE_SOCIALE_REÇU), dtpAdoptions(dtpPSYCHOLOGUE_INTERVENTION), dtpAdoptions(dtpPSYCHOLOGUE_REÇU)
    dtpAdoptions(dtpSURVEILLANCE) = "": dtpAdoptions(dtpENVOI_TGI) = "": dtpAdoptions(dtpASSISTANTE_SOCIALE_INTERVENTION) = "": dtpAdoptions(dtpASSISTANTE_SOCIALE_REÇU) = "": dtpAdoptions(dtpPSYCHOLOGUE_INTERVENTION) = "": dtpAdoptions(dtpPSYCHOLOGUE_REÇU) = ""
    
'   Lire les références du placement
    Set prstPlacement = modDAO.rstSQLExtraction(daoPLACEMENT_RÉFÉRENCES, True, gstrCONSULTATION_LIMITÉE, mlngNuméroPlacement)
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
        If modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT Then Err.Raise mesENREGISTREMENT_SUPPRIMÉ
'       Interdire la validation
        Implémenter "<interdiction de valider>", "Vrai"
        Exit Sub
    End If
    
    With prstPlacement
    
        mlngNuméroDemande = ![DEMA_NUM]
        
'       Si la date de réception du dossier n'est pas saisie
        If IsNull(![DEMA_DREDOS]) Then
            Err.Raise mesMESSAGE_020
'           Interdire la validation
            Implémenter "<interdiction de valider>", "Vrai"
            Exit Sub
'       Sinon
        Else
            mdtmRéceptionDossier = ![DEMA_DREDOS]
        End If
    
    End With

'   Vérifier l'existence de surveillances
    Set mrstSurveillances = modDAO.rstSQLExtraction(daoSURVEILLANCES, True, gstrCONSULTATION_LIMITÉE, dbOpenDynaset, mlngNuméroPlacement)
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
'       Interdire la validation
        Implémenter "<interdiction de valider>", "Vrai"
        Exit Sub
    End If
    
'   S'il s'agit d'une création
    If mlngTraitement = comCRÉATION Then
    
'       S'il existe une ou plusieurs surveillances
        If Not modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT Then
            With mrstSurveillances
'               Proposer la même date de demande de surveillance et les mêmes intervenants que lors de la dernière surveillance
                dtpAdoptions(dtpSURVEILLANCE) = ![SUPL_DDSUR]: dtpAdoptions(dtpSURVEILLANCE) = ""
                mvarAssistanteSocialeNuméro = ![TRSO_NUM_AssistanteSociale]: modAdoptions.AfficherIdentitéIntervenant CInt(mvarAssistanteSocialeNuméro), txtAdoptions(txtASSISTANTE_SOCIALE), gstrCONSULTATION_LIMITÉE
                mvarPsychologueNuméro = ![TRSO_NUM_Psychologue]: modAdoptions.AfficherIdentitéIntervenant CInt(mvarPsychologueNuméro), txtAdoptions(txtPSYCHOLOGUE), gstrCONSULTATION_LIMITÉE
            End With
        End If
    
'   Sinon
    Else
    
        mdtmSurveillance = pvarSurveillance: mintNuméroRapport = pvarNuméroRapport
        
'       Si la surveillance à mettre à jour ou à supprimer n'existe plus, sortir
        If Not booPositionnementSurveillanceEffectué(True) Then Exit Sub
        
        With mrstSurveillances
'           Vérifier la cohérence entre les informations présentes sur la surveillance et celles de l'historique des intervenants
            modAdoptions.VérifierCohérenceAvecHistoriqueIntervenants mlngNuméroDemande, mintNuméroRapport, adpSURVEILLANCE_PLACEMENT, ![TRSO_NUM_Psychologue], ![SUPL_DDSUR], ![SUPL_dtmPsychologuePrevu], ![SUPL_dtmPsychologueRecu], ![TRSO_NUM_AssistanteSociale], ![SUPL_DDSUR], ![SUPL_dtmAssistanteSocialePrevu], ![SUPL_dtmAssistanteSocialeRecu]
        End With
                    
    End If
    
'   Activer/Désactiver les contrôles se rapportant à la surveillance
    dtpAdoptions_Change dtpSURVEILLANCE
    
    Select Case mlngTraitement
    
'       Création de la surveillance
        Case comCRÉATION
        
            pstrMessages(0) = "Création"
        
'       Mise à jour de la surveillance
        Case comMODIFICATION
        
            pstrMessages(0) = "Mise à jour"
        
'           Désactiver la date de la demande de surveillance
            modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpSURVEILLANCE), True, True
            
'       Suppression de la surveillance
        Case comSUPPRESSION
        
            pstrMessages(0) = "Suppression"
        
'           Désactiver la saisie
            modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpSURVEILLANCE), True, True
            modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpENVOI_TGI), True, True
            modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdASSISTANTE_SOCIALE), True
            modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpASSISTANTE_SOCIALE_INTERVENTION), True, True
            modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpASSISTANTE_SOCIALE_REÇU), True, True
            modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdPSYCHOLOGUE), True
            modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpPSYCHOLOGUE_INTERVENTION), True, True
            modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpPSYCHOLOGUE_REÇU), True, True
            
    End Select
            
    Exit Sub
    
Err_Initialiser:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                 pintTraitement = mesINTERROMPRE
        Case mesENREGISTREMENT_SUPPRIMÉ:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, pstrMessages(0), "Le placement")
        Case mesMESSAGE_020:                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
        Case Else:                          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


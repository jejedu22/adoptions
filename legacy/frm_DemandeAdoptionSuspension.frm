VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_DemandeAdoptionSuspension 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instruction de la demande"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraAdoptions 
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
      Height          =   2535
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   10095
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
         TabIndex        =   30
         Top             =   1560
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
         Index           =   0
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1080
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
         Index           =   1
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1560
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
         Index           =   1
         Left            =   1740
         TabIndex        =   24
         Top             =   1080
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTP_DateReprise 
         Height          =   285
         Left            =   8280
         TabIndex        =   18
         ToolTipText     =   "Date de la commission"
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58720259
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   0
         Left            =   4695
         TabIndex        =   26
         ToolTipText     =   "Date de demande du rapport à l'assistante sociale"
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58720259
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   3
         Left            =   4695
         TabIndex        =   32
         ToolTipText     =   "Date de demande du rapport au psychologue"
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58720259
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   1
         Left            =   6510
         TabIndex        =   27
         ToolTipText     =   "Date de réception du rapport de l'assistante sociale"
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58720259
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   4
         Left            =   6510
         TabIndex        =   33
         ToolTipText     =   "Date de réception du rapport du psychologue"
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58720259
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   2
         Left            =   8325
         TabIndex        =   28
         ToolTipText     =   "Date prévue de réception du rapport de l'AS"
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58720259
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   5
         Left            =   8325
         TabIndex        =   34
         ToolTipText     =   "Date prévue de réception du rapport du psychologue"
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58720259
         CurrentDate     =   2958465
      End
      Begin VB.Label Lbl_DateReprise 
         Caption         =   " Décidée le"
         Height          =   195
         Left            =   7380
         TabIndex        =   17
         Top             =   45
         Width           =   1005
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         X1              =   240
         X2              =   9840
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         X1              =   240
         X2              =   9840
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Assistante sociale"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   23
         Top             =   1110
         Width           =   1260
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         X1              =   240
         X2              =   9840
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Psychologue"
         Height          =   195
         Index           =   8
         Left            =   585
         TabIndex        =   29
         Top             =   1590
         Width           =   915
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF0000&
         X1              =   1620
         X2              =   1620
         Y1              =   720
         Y2              =   1920
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   2  'Center
         Caption         =   "Nom"
         Height          =   255
         Index           =   3
         Left            =   1740
         TabIndex        =   19
         Top             =   705
         Width           =   2295
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FF0000&
         X1              =   4575
         X2              =   4575
         Y1              =   720
         Y2              =   1920
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FF0000&
         X1              =   6390
         X2              =   6390
         Y1              =   720
         Y2              =   1920
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   2  'Center
         Caption         =   "Rapport demandé le"
         Height          =   255
         Index           =   4
         Left            =   4695
         TabIndex        =   20
         Top             =   705
         Width           =   1575
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   2  'Center
         Caption         =   "Rapport reçu le"
         Height          =   240
         Index           =   5
         Left            =   6510
         TabIndex        =   21
         Top             =   705
         Width           =   1575
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FF0000&
         X1              =   8205
         X2              =   8205
         Y1              =   720
         Y2              =   1920
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   2  'Center
         Caption         =   "Réception prévue le"
         Height          =   240
         Index           =   6
         Left            =   8325
         TabIndex        =   22
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Reprise"
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
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   660
      End
   End
   Begin VB.Frame fraAdoptions 
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
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   10095
      Begin VB.TextBox Txt_Motif 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   37
         ToolTipText     =   "Motif de suspension de la demande"
         Top             =   930
         Width           =   6540
      End
      Begin MSComCtl2.DTPicker DTP_DateSuspension 
         Height          =   285
         Left            =   8280
         TabIndex        =   14
         ToolTipText     =   "Date de suspension de la demande"
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58720259
         CurrentDate     =   2958465
      End
      Begin VB.Label lbl_Motif 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Motif"
         Height          =   195
         Left            =   4823
         TabIndex        =   38
         Top             =   570
         Width           =   375
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Suspension"
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
         TabIndex        =   36
         Top             =   0
         Width           =   990
      End
      Begin VB.Label Lbl_DateSuspension 
         Caption         =   " Décidée le"
         Height          =   195
         Left            =   7380
         TabIndex        =   13
         Top             =   45
         Width           =   1005
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   945
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   615
         Left            =   9240
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
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Annuler"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
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
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   10095
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
         Left            =   3705
         MaxLength       =   50
         TabIndex        =   7
         ToolTipText     =   "Motif du refus suite au recours administratif"
         Top             =   360
         Width           =   6150
      End
      Begin VB.ComboBox Cmb_NumDemande 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox Cmb_NumDossier 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTP_DateDemande 
         Height          =   285
         Left            =   3720
         TabIndex        =   11
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
         Format          =   58720259
         CurrentDate     =   2958465
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Demandeur"
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   6
         Top             =   405
         Width           =   825
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
      Begin VB.Label Lbl_DateDemande 
         AutoSize        =   -1  'True
         Caption         =   "Déposée le"
         Height          =   195
         Left            =   2775
         TabIndex        =   10
         Top             =   765
         Width           =   810
      End
   End
   Begin MSComctlLib.StatusBar stb_DemandesSuspension 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   7035
      Width           =   10290
      _ExtentX        =   18150
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
            Object.Width           =   11924
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_DemandeAdoptionSuspension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mfrm_TS As frm_TS
Attribute mfrm_TS.VB_VarHelpID = -1

Private Enum IndexCommandButton         'Index des contrôles CommandButton
    cmdASSISTANTE_SOCIALE = 0
    cmdPSYCHOLOGUE
End Enum
Private Enum IndexDTPicker              'Index des contrôles DTPicker
    dtpASSISTANTE_SOCIALE_DEMANDÉ = 0
    dtpASSISTANTE_SOCIALE_REÇU
    dtpASSISTANTE_SOCIALE_PRÉVU
    dtpPSYCHOLOGUE_DEMANDÉ
    dtpPSYCHOLOGUE_REÇU
    dtpPSYCHOLOGUE_PRÉVU
End Enum
Private Enum IndexFrame                 'Index des contrôles Frame
    fraSUSPENSION = 0
    fraREPRISE
End Enum
Private Enum IndexLabel                 'Index des contrôles Label
    lblDEMANDEUR = 0
    lblSUSPENSION
    lblREPRISE
End Enum
Private Enum IndexTextBox               'Index des contrôles TextBox
    txtDEMANDEUR = 0
    txtASSISTANTE_SOCIALE
    txtPSYCHOLOGUE
End Enum

Private mvarPsychologueNuméro As Variant, mvarAssistanteSocialeNuméro As Variant
Private mrstDemande As Recordset
Private mlngIntervenant As IndexCommandButton
Private mbooInitialisationAvecDemandeDéfinie As Boolean, mbooSuspensionAutorisée As Boolean


Private Sub Cmb_NumDemande_Click()
Const pkstrNomProcédure As String = "Cmb_NumDemande_Click"
On Error GoTo Err_Cmb_NumDemande_Click
        
'   Implémenter le titre
    Me.Caption = "Instruction de la demande n° " & Cmb_NumDemande.Text & " / dossier n° " + Cmb_NumDossier
    
'   Implémenter la date de dépôt de la demande
    Implémenter "Déposée le", "True"
    
'   Activer les frames
    Implémenter "<désactivation des frames>", "Faux"
    
'   Lire la demande
    Set mrstDemande = modDAO.rstSQLExtraction(daoEXTRACTION_028, True, gstrCONSULTATION_LIMITÉE & gstrSÉPARATEUR_1 & gstrDEMANDE_SUPPRIMÉE, CLng(Cmb_NumDemande))
'   Déterminer si la validation est autorisée
    Implémenter "<interdiction de valider>", CStr((Not modDAO.lngDernièreErreurDAO() = daoAUCUNE))
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then Exit Sub

'   Positionner les contrôles
    With mrstDemande
        
        DTP_DateDemande = modCommun.varValeurFormatéeSiNull(![DEMA_DATE], vbDate, Null)
    
        mbooSuspensionAutorisée = Not (modCommun.varValeurFormatéeSiNull(![DEMA_DSUSP], vbDate, #12:00:00 AM#) > modCommun.varValeurFormatéeSiNull(![DEMA_DREPR], vbDate, #12:00:00 AM#))
        
        lblAdoptions(lblSUSPENSION).ForeColor = IIf(mbooSuspensionAutorisée, vbBlue, vbGrayText)
        lblAdoptions(lblREPRISE).ForeColor = IIf(Not mbooSuspensionAutorisée, vbBlue, vbGrayText)
        DTP_DateSuspension = IIf(mbooSuspensionAutorisée, Date, #12/31/9999#)
        DTP_DateReprise = IIf(Not mbooSuspensionAutorisée, Date, #12/31/9999#)
        
'       Si la suspension est autorisée
        If mbooSuspensionAutorisée Then
            DTP_DateSuspension = ""
            Txt_Motif = ""
            DTP_DateReprise = modCommun.varValeurFormatéeSiNull(![DEMA_DREPR], vbDate, Null)
'       Sinon
        Else
            DTP_DateSuspension = modCommun.varValeurFormatéeSiNull(![DEMA_DSUSP], vbDate, Null)
            Txt_Motif = modCommun.varValeurFormatéeSiNull(![DEMA_MSUSP], vbString)
            DTP_DateReprise = ""
        End If
        
        modCommun.DéfinirVerrouillageContrôle DTP_DateSuspension, Not mbooSuspensionAutorisée, Not mbooSuspensionAutorisée
        modCommun.DéfinirVerrouillageContrôle DTP_DateReprise, mbooSuspensionAutorisée, Not IsNull(![DEMA_DREPR])
        
    End With
    
'   Implémenter le titre
    Me.Caption = Me.Caption & " - " & IIf(mbooSuspensionAutorisée, "Suspension", "Reprise")
    
'   Activer/Désactiver l'ensemble des contrôles se rapportant à la suspension selon que la date de suspension est saisie ou non
    DTP_DateSuspension_Change
'   Activer/Désactiver l'ensemble des contrôles se rapportant à la reprise selon que la date de reprise est saisie ou non
    DTP_DateReprise_Change

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
    modAdoptions.AfficherIdentitéDemandeur Cmb_NumDossier, txtAdoptions(txtDEMANDEUR), gstrCONSULTATION_LIMITÉE, cmd_Valider

'   Implémenter le titre
    Me.Caption = "Instruction de la demande " & "/ dossier n° " + Cmb_NumDossier
    
'   Implémenter la date de dépôt de la demande
    Implémenter "Déposée le", "False"
    
'   Désactiver les frames
    Implémenter "<désactivation des frames>", "Vrai"

'   Si la demande n'est pas définie à l'initialisation de la fenêtre
    If Not mbooInitialisationAvecDemandeDéfinie Then
'       Lire les numéros des demandes liées au dossier
        modAdoptions.ImplémenterListe Cmb_NumDemande, "Demandes", gstrCONSULTATION_LIMITÉE, cmd_Valider, Cmb_NumDossier
'       S'il n'en existe qu'une, la sélectionner
        If Cmb_NumDemande.ListCount = 1 Then Cmb_NumDemande.ListIndex = 0
    End If

    If Cmb_NumDemande.Enabled Then Cmb_NumDemande.SetFocus
    
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
Const pkstrListeMessages As String = gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_1 & "Validation refusée: Un rapport <de l'intervenant> a déjà été demandé pour cette date."
Dim pvarNuméro As Variant
Dim pctlFocus As Control
Dim pcmdIntervenant As CommandButton
Dim pdtpDemandé As DTPicker, pdtpReçu As DTPicker, pdtpPrévu As DTPicker
Dim pstrMessages(1) As String, pstrChaîne As String, pstrListeMessages As String
Dim plngParamètreÉchéancier As DAOParamètresÉchéancier
Dim pintItération As Integer
Dim pbooSupprimerReprise As Boolean
On Error GoTo Err_cmd_Valider_Click
    
'   Forcer si nécessaire l'actualisation de la valeur d'un contrôle DTPicker
    modCommun.GérerActualisationDTPicker Me, cmd_Valider
    
'   Refuser la validation si le numéro de la demande n'est pas sélectionné
    If Cmb_NumDemande = "" Then
        Set pctlFocus = Cmb_NumDemande: pstrMessages(0) = "du numéro de la demande"
        Err.Raise mesSÉLECTION_MANQUANTE
        Exit Sub
    End If
    
    Select Case mbooSuspensionAutorisée
'       Suspension
        Case True:  Set pctlFocus = DTP_DateSuspension: pstrMessages(0) = "de la date de la suspension"
'       Reprise
        Case False: Set pctlFocus = DTP_DateReprise:    pstrMessages(0) = "de la date de la reprise"
    End Select
    
'   Refuser la validation si la date de la suspension / reprise n'est pas définie
    If IsNull(pctlFocus) Then Err.Raise mesSAISIE_MANQUANTE: Exit Sub

'   Ouvrir une transaction
    modDAO.GérerTransaction daoOUVERTURE
    
    mrstDemande.Edit
    
'   S'il s'agit d'une suspension
    If mbooSuspensionAutorisée Then
    
'       Refuser la validation si la date de suspension est antérieure à la date de dépôt de la demande
        If DTP_DateSuspension < modCommun.varValeurFormatéeSiNull(DTP_DateDemande, vbDate, #12:00:00 AM#) Then
            Set pctlFocus = DTP_DateSuspension: pstrMessages(0) = "de suspension": pstrMessages(1) = "de dépôt de la demande"
            Err.Raise mesDATE_ANTÉRIEURE
            modDAO.GérerTransaction daoANNULATION: Exit Sub
        End If

'       Si la date de suspension est antérieure ou égale à la date de reprise précédente
        If Not DTP_DateSuspension > modCommun.varValeurFormatéeSiNull(DTP_DateReprise, vbDate, #12:00:00 AM#) Then
'           Demander confirmation
            Select Case modMessages.lngRéponseUtilisateur(mesQUESTION_010)
'               L'utilisateur confirme la validation
                Case vbYes
                    pbooSupprimerReprise = True
'               L'utilisateur décide l'abandon du traitement
                Case vbNo
                    modDAO.GérerTransaction daoANNULATION: Exit Sub
            End Select
        End If

        mrstDemande![DEMA_DSUSP] = DTP_DateSuspension
        mrstDemande![DEMA_MSUSP] = modCommun.varChaîneFormatée(Txt_Motif, vbString, Null)
        If pbooSupprimerReprise Then mrstDemande![DEMA_DREPR] = Null
        
'   Sinon
    Else
    
'       Refuser la validation si la date de reprise est antérieure à la date de suspension
        If DTP_DateReprise < DTP_DateSuspension Then
            Set pctlFocus = DTP_DateReprise: pstrMessages(0) = "de reprise": pstrMessages(1) = "de suspension"
            Err.Raise mesDATE_ANTÉRIEURE
            modDAO.GérerTransaction daoANNULATION: Exit Sub
        End If

'       Pour chaque intervenant
        For pintItération = 1 To 2
    
            Select Case pintItération
'               Assistante sociale
                Case 1: pvarNuméro = mvarAssistanteSocialeNuméro:   Set pcmdIntervenant = cmdAdoptions(cmdASSISTANTE_SOCIALE):  Set pdtpDemandé = dtpAdoptions(dtpASSISTANTE_SOCIALE_DEMANDÉ):  Set pdtpReçu = dtpAdoptions(dtpASSISTANTE_SOCIALE_REÇU):    Set pdtpPrévu = dtpAdoptions(dtpASSISTANTE_SOCIALE_PRÉVU):  plngParamètreÉchéancier = daoRÉCEPTION_DOSSIER__RÉPONSE_TRAVAILLEUR_SOCIAL: pstrChaîne = "de l'assistante sociale": pstrListeMessages = Replace(pkstrListeMessages, "<de l'intervenant>", "de l'assistante sociale")
'               Psychologue
                Case 2: pvarNuméro = mvarPsychologueNuméro:         Set pcmdIntervenant = cmdAdoptions(cmdPSYCHOLOGUE):         Set pdtpDemandé = dtpAdoptions(dtpPSYCHOLOGUE_DEMANDÉ):         Set pdtpReçu = dtpAdoptions(dtpPSYCHOLOGUE_REÇU):           Set pdtpPrévu = dtpAdoptions(dtpPSYCHOLOGUE_PRÉVU):         plngParamètreÉchéancier = daoRÉCEPTION_DOSSIER__RÉPONSE_PSYCHOLOGUE:        pstrChaîne = "du psychologue":          pstrListeMessages = Replace(pkstrListeMessages, "<de l'intervenant>", "du psychologue")
            End Select
            
'           Si un rapport de l'intervenant est demandé
            If Not IsNull(pdtpDemandé) Then
'               Si la date prévue pour sa réception n'est pas définie
                If IsNull(pdtpPrévu) Then
'                   La calculer
                    pdtpPrévu = DateAdd("m", modDAO.intParamètreÉchéancier(plngParamètreÉchéancier), pdtpDemandé)
                End If
            End If
        
'           Refuser la validation si les informations relatives à l'intervenant sont incohérentes
            If modAdoptions.booSaisieIncohérenteIntervenant(pdtpDemandé, DTP_DateReprise, pvarNuméro, pcmdIntervenant, pdtpDemandé, pdtpReçu, pstrChaîne, "de demande du rapport", "de la reprise", pdtpPrévu:=pdtpPrévu) Then
                modDAO.GérerTransaction daoANNULATION: Exit Sub
            End If
            
'           Si l'intervenant est identifié
            If Not IsNull(pdtpDemandé) Then
'               Créer l'historique
                modDAO.SQLAjouter daoADD_HISTORIQUE_INTERVENANT, True, pstrListeMessages, CLng(Cmb_NumDemande), adpREPRISE, (pintItération = 2), pdtpDemandé, 1, pdtpPrévu, pdtpReçu, pvarNuméro
                If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
            End If
                    
        Next pintItération
        
        mrstDemande![DEMA_DSUSP] = Null
        mrstDemande![DEMA_MSUSP] = Null
        mrstDemande![DEMA_DREPR] = DTP_DateReprise
        
    End If
    
'   Valider l'opération
    modDAO.ValiderOpérationSurRecordset mrstDemande
    Select Case modDAO.lngDernièreErreurDAO()
'       L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
        Case daoENREGISTREMENT_VERROUILLÉ
            modDAO.GérerTransaction daoANNULATION: Exit Sub
'       L'opération a été refusée en raison de la modification par un autre utilisateur des données originales
        Case dbRecordModified
'           Interdire la validation
            Implémenter "<interdiction de valider>", "Vrai"
            modDAO.GérerTransaction daoANNULATION: Exit Sub
'       L'opération a été refusée en raison de la suppression par un autre utilisateur des données originales
        Case dbRecordDeleted
'           Interdire la validation
            Implémenter "<interdiction de valider>", "Vrai"
            modDAO.GérerTransaction daoANNULATION: Exit Sub
'       L'opération a généré un conflit inattendu
        Case Is <> daoAUCUNE
            modDAO.GérerTransaction daoANNULATION
            Err.Raise mesCONFLIT_ACCÈS
    End Select

'   Actualiser les données relatives aux intervenants
    modAdoptions.ActualiserDonnéesIntervenants CLng(Cmb_NumDemande), gstrVALIDATION_ANNULÉE
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
        modDAO.GérerTransaction daoANNULATION: Exit Sub
    End If

'   Valider la transaction
    modDAO.GérerTransaction daoVALIDATION

    Unload Me
 
    Exit Sub
  
Err_cmd_Valider_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:             pintTraitement = mesINTERROMPRE
        Case mesCONFLIT_ACCÈS:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, modDAO.lngDernièreErreurDAO())
        Case mesSÉLECTION_MANQUANTE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, pstrMessages(0))
        Case mesSAISIE_MANQUANTE:       pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, pstrMessages(0))
        Case mesDATE_ANTÉRIEURE:        pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, pstrMessages(0), pstrMessages(1))
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
        .Initialiser False, True, True
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

Private Sub DTP_DateReprise_Change()
Const pkstrNomProcédure As String = "DTP_DateReprise_Change"
On Error GoTo Err_DTP_DateReprise_Change

'   Activer/Désactiver les contrôles se rapportant à la reprise
    Implémenter "<activation de la reprise>,<contrôles se rapportant à la reprise>"
    
    Exit Sub
    
Err_DTP_DateReprise_Change:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub DTP_DateReprise_GotFocus()
    stb_DemandesSuspension.SimpleText = "Entrez la date de reprise de l'instruction de la demande"
End Sub

Private Sub DTP_DateSuspension_Change()
Const pkstrNomProcédure As String = "DTP_DateSuspension_Change"
On Error GoTo Err_DTP_DateSuspension_Change

'   Activer/Désactiver les contrôles se rapportant à la suspension
    Implémenter "<activation de la suspension>"
    
    Exit Sub
    
Err_DTP_DateSuspension_Change:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub DTP_DateSuspension_GotFocus()
    stb_DemandesSuspension.SimpleText = "Entrez la date de suspension de la demande"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const pkstrNomProcédure As String = "Form_Unload"
On Error GoTo Err_Form_Unload

    Set mrstDemande = Nothing
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

Private Sub Txt_Motif_Gotfocus()
    stb_DemandesSuspension.SimpleText = "Entrez le motif de suspension de la demande"
End Sub


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
        
'           Date de dépôt de la demande
            Case "Déposée le"
        
'               Positionner la visibilité
                pstrParamètres() = Split(pstrListesParamètres(pintItération), gstrSÉPARATEUR_2)
                DTP_DateDemande.Visible = CBool(pstrParamètres(0)): Lbl_DateDemande.Visible = CBool(pstrParamètres(0))
        
'           Gestion de l'activation des contrôles se rapportant à la reprise
            Case "<activation de la reprise>"
        
                pbooVerrouillé = IsNull(DTP_DateReprise) Or mbooSuspensionAutorisée
                
                modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdASSISTANTE_SOCIALE), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpASSISTANTE_SOCIALE_DEMANDÉ), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpASSISTANTE_SOCIALE_REÇU), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpASSISTANTE_SOCIALE_PRÉVU), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdPSYCHOLOGUE), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpPSYCHOLOGUE_DEMANDÉ), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpPSYCHOLOGUE_REÇU), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpPSYCHOLOGUE_PRÉVU), pbooVerrouillé
                
'           Gestion de l'activation des contrôles se rapportant à la suspension
            Case "<activation de la suspension>"
        
                pbooVerrouillé = IsNull(DTP_DateSuspension) Or Not mbooSuspensionAutorisée
                
                modCommun.DéfinirVerrouillageContrôle Txt_Motif, pbooVerrouillé, Not mbooSuspensionAutorisée
                If IsNull(DTP_DateSuspension) Then Txt_Motif = ""
                
'           Gestion de l'activation des frames
            Case "<désactivation des frames>"
        
                pstrParamètres() = Split(pstrListesParamètres(pintItération), gstrSÉPARATEUR_2)
                pbooVerrouillé = CBool(pstrParamètres(0))
                
                modCommun.DéfinirVerrouillageContrôle fraAdoptions(fraSUSPENSION), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle fraAdoptions(fraREPRISE), pbooVerrouillé
                
'           Dates de suspension et de reprise
            Case "<dates de suspension et de reprise>"
        
'               Désélectionner les dates
                DTP_DateSuspension = "": DTP_DateReprise = ""
                
'               Implémenter les contrôles se rapportant à la reprise
                Implémenter "<contrôles se rapportant à la reprise>"
                
'           Contrôles se rapportant à la reprise
            Case "<contrôles se rapportant à la reprise>"
        
                If IsNull(DTP_DateReprise) Or mbooSuspensionAutorisée Then
                    mvarPsychologueNuméro = Null: mvarAssistanteSocialeNuméro = Null
                    txtAdoptions(txtASSISTANTE_SOCIALE) = "": txtAdoptions(txtPSYCHOLOGUE) = ""
                    dtpAdoptions(dtpASSISTANTE_SOCIALE_DEMANDÉ) = #12/31/9999#: dtpAdoptions(dtpASSISTANTE_SOCIALE_REÇU) = #12/31/9999#: dtpAdoptions(dtpASSISTANTE_SOCIALE_PRÉVU) = #12/31/9999#: dtpAdoptions(dtpPSYCHOLOGUE_DEMANDÉ) = #12/31/9999#: dtpAdoptions(dtpPSYCHOLOGUE_REÇU) = #12/31/9999#: dtpAdoptions(dtpPSYCHOLOGUE_PRÉVU) = #12/31/9999#
                Else
                    modCommun.PositionnerDateJour dtpAdoptions(dtpASSISTANTE_SOCIALE_DEMANDÉ), dtpAdoptions(dtpASSISTANTE_SOCIALE_REÇU), dtpAdoptions(dtpASSISTANTE_SOCIALE_PRÉVU), dtpAdoptions(dtpPSYCHOLOGUE_DEMANDÉ), dtpAdoptions(dtpPSYCHOLOGUE_REÇU), dtpAdoptions(dtpPSYCHOLOGUE_PRÉVU)
                End If
                
'               Désélectionner les dates
                dtpAdoptions(dtpASSISTANTE_SOCIALE_DEMANDÉ) = "": dtpAdoptions(dtpASSISTANTE_SOCIALE_REÇU) = "": dtpAdoptions(dtpASSISTANTE_SOCIALE_PRÉVU) = "": dtpAdoptions(dtpPSYCHOLOGUE_DEMANDÉ) = "": dtpAdoptions(dtpPSYCHOLOGUE_REÇU) = "": dtpAdoptions(dtpPSYCHOLOGUE_PRÉVU) = ""
                
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
'   pvarNuméroDemande:  numéro de demande [facultatif]

Public Sub Initialiser(Optional pvarNuméroDemande As Variant)
Const pkstrNomProcédure As String = "Initialiser"
Dim pstrNuméroDossier As String
On Error GoTo Err_Initialiser
    
    mbooInitialisationAvecDemandeDéfinie = Not IsMissing(pvarNuméroDemande)
    
'   Centrer la feuille
    modAdoptions.CentrerFeuille Me
        
'   Implémenter les images
    cmd_Quitter.Picture = Menu.picImage(adp32x32, adpANNULER)
    cmd_Valider.Picture = Menu.picImage(adp32x32, adpVALIDER)
    cmdAdoptions(cmdASSISTANTE_SOCIALE).Picture = Menu.picImage(adp16x16, adpINTERVENANTS)
    cmdAdoptions(cmdPSYCHOLOGUE).Picture = Menu.picImage(adp16x16, adpINTERVENANTS)
'
'   Désactiver certains contrôles
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtDEMANDEUR), True
    modCommun.DéfinirVerrouillageContrôle DTP_DateDemande, True
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtASSISTANTE_SOCIALE), True
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtPSYCHOLOGUE), True

    Set mfrm_TS = New frm_TS

'   Implémenter la date de dépôt de la demande et les dates de suspension et de reprise
    Implémenter "Déposée le,<dates de suspension et de reprise>", "Faux" & gstrSÉPARATEUR_1 & ""
    
'   Si la demande est définie
    If mbooInitialisationAvecDemandeDéfinie Then
    
        pstrNuméroDossier = modAdoptions.txtNuméroDossier(CLng(pvarNuméroDemande), cmd_Valider)
        If pstrNuméroDossier = "" Then Exit Sub
        
        modCommun.DéfinirVerrouillageContrôle Cmb_NumDossier, True, True
        modCommun.DéfinirVerrouillageContrôle Cmb_NumDemande, True, True
        modCommun.ImplémenterListe Cmb_NumDossier, pstrNuméroDossier
        modCommun.ImplémenterListe Cmb_NumDemande, CStr(pvarNuméroDemande)
        Cmb_NumDossier.ListIndex = modCommun.varRecherche(Cmb_NumDossier, True, pstrNuméroDossier)
        Cmb_NumDemande.ListIndex = modCommun.varRecherche(Cmb_NumDemande, True, CStr(pvarNuméroDemande))
        
'   Sinon
    Else
    
        modAdoptions.ImplémenterListe Cmb_NumDossier, "Dossiers", gstrCONSULTATION_LIMITÉE, cmd_Valider
        
'       Désactiver les frames
        Implémenter "<désactivation des frames>", "Vrai"
    
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


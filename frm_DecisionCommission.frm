VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_DecisionCommission 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Décision de la commission"
   ClientHeight    =   8865
   ClientLeft      =   75
   ClientTop       =   345
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   8760
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
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   8535
      Begin VB.ComboBox cboAdoptions 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Événement ayant conduit à délivrer l'agrément"
         Top             =   360
         Visible         =   0   'False
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   15
         ToolTipText     =   "Date de délivrance de l'agrément"
         Top             =   360
         Visible         =   0   'False
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
         Format          =   58916867
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   19
         ToolTipText     =   "Date du retrait"
         Top             =   720
         Visible         =   0   'False
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
         Format          =   58916867
         CurrentDate     =   2958465
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Agrément"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   810
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Retrait décidé le"
         Height          =   195
         Index           =   3
         Left            =   1005
         TabIndex        =   18
         Top             =   765
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Suite à"
         Height          =   195
         Index           =   2
         Left            =   4185
         TabIndex        =   16
         Top             =   405
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Délivré le"
         Height          =   195
         Index           =   1
         Left            =   1500
         TabIndex        =   14
         Top             =   405
         Visible         =   0   'False
         Width           =   660
      End
   End
   Begin VB.Frame fra_enfants 
      Caption         =   "Souhaits"
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
      Left            =   120
      TabIndex        =   35
      Top             =   6000
      Width           =   8535
      Begin VB.CheckBox chkAdoptions 
         Alignment       =   1  'Right Justify
         Caption         =   "Fratrie"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   36
         ToolTipText     =   "Fratrie autorisée ?"
         Top             =   450
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.ComboBox Cmb_OrigEt 
         Height          =   315
         Index           =   3
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   41
         ToolTipText     =   "Origine ethnique"
         Top             =   1920
         Width           =   3615
      End
      Begin VB.ComboBox Cmb_OrigEt 
         Height          =   315
         Index           =   2
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   40
         ToolTipText     =   "Origine ethnique"
         Top             =   1620
         Width           =   3615
      End
      Begin VB.ComboBox Cmb_OrigEt 
         Height          =   315
         Index           =   1
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   39
         ToolTipText     =   "Origine ethnique"
         Top             =   1320
         Width           =   3615
      End
      Begin VB.ComboBox Cmb_OrigEt 
         Height          =   315
         Index           =   0
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   38
         ToolTipText     =   "Origine ethnique"
         Top             =   1020
         Width           =   3615
      End
      Begin VB.CommandButton cmd_CreerIndividu 
         Height          =   375
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Créer un enfant"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmd_ModifierEnfant 
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Modifier l'enfant"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmd_SupprimerEnfant 
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Supprimer l'enfant"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.TextBox Txt_NombreEnfants 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   195
         Left            =   7200
         TabIndex        =   45
         Text            =   "0 enfant(s)"
         ToolTipText     =   "Nombre d'enfants demandés"
         Top             =   450
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwAdoptions 
         Height          =   1200
         Index           =   0
         Left            =   4800
         TabIndex        =   46
         ToolTipText     =   "Liste des enfants"
         Top             =   1020
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   2117
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Numéro"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Sexe"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Handicap"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Tranche d'âges"
            Object.Width           =   2469
         EndProperty
      End
      Begin VB.Label Lbl_origine 
         Caption         =   "Origines ethniques"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   945
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   615
         Left            =   6960
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
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Valider"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame fra_commission 
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
      Height          =   2055
      Left            =   120
      TabIndex        =   20
      Top             =   3840
      Width           =   8535
      Begin VB.CommandButton cmdAdoptions 
         Height          =   375
         Index           =   0
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Visualiser l'historique des décisions"
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.ComboBox Cmb_Deci 
         Height          =   315
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   31
         ToolTipText     =   "Décision finale"
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox Cmb_Avis 
         Height          =   315
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   29
         ToolTipText     =   "Avis de la commission"
         Top             =   360
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTP_DateDeci 
         Height          =   285
         Left            =   2280
         TabIndex        =   25
         ToolTipText     =   "Date de signature de la décision"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         DateIsNull      =   -1  'True
         Format          =   58916867
         CurrentDate     =   36328
      End
      Begin MSComCtl2.DTPicker DTP_DateReCom 
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         ToolTipText     =   "Date de réunion de la commission"
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         DateIsNull      =   -1  'True
         Format          =   58916867
         CurrentDate     =   36328
      End
      Begin VB.TextBox txt_motifrefus 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   33
         ToolTipText     =   "Motif du refus ou du retrait"
         Top             =   1485
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker DTP_PreDeci 
         Height          =   285
         Left            =   2280
         TabIndex        =   27
         ToolTipText     =   "Date prévue pour la signature de la décision"
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         DateIsNull      =   -1  'True
         Format          =   58916867
         CurrentDate     =   36328
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Prochaine / Dernière commission"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   2820
      End
      Begin VB.Label Lbl_dateprevuecomm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date prévue de la décision"
         Height          =   195
         Left            =   270
         TabIndex        =   26
         Top             =   1125
         Width           =   1905
      End
      Begin VB.Label lbl_refus 
         AutoSize        =   -1  'True
         Caption         =   "Motif du refus/retrait"
         Height          =   195
         Left            =   750
         TabIndex        =   32
         Top             =   1515
         Width           =   1425
      End
      Begin VB.Label lbl_deci 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Décision finale"
         Height          =   195
         Left            =   5535
         TabIndex        =   30
         Top             =   765
         Width           =   1035
      End
      Begin VB.Label lbl_aviscom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Avis de la commission"
         Height          =   195
         Left            =   5025
         TabIndex        =   28
         Top             =   405
         Width           =   1545
      End
      Begin VB.Label lbl_datedeci 
         AutoSize        =   -1  'True
         Caption         =   "Date de la signature"
         Height          =   195
         Left            =   750
         TabIndex        =   24
         Top             =   765
         Width           =   1425
      End
      Begin VB.Label lbl_datecom 
         AutoSize        =   -1  'True
         Caption         =   "Date de la commission"
         Height          =   195
         Left            =   585
         TabIndex        =   22
         Top             =   405
         Width           =   1590
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
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   8535
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
         ToolTipText     =   "Demandeur"
         Top             =   360
         Width           =   4590
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
      Begin VB.ComboBox Cmb_NumDemande 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Numéro de la demande"
         Top             =   720
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTP_DateDemande 
         Height          =   285
         Left            =   3720
         TabIndex        =   11
         ToolTipText     =   "Date de dépôt de la demande"
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
         Format          =   58916867
         CurrentDate     =   2958465
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Demandeur"
         Height          =   195
         Index           =   5
         Left            =   2760
         TabIndex        =   6
         Top             =   405
         Width           =   825
      End
      Begin VB.Label Lbl_DateDemande 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Déposée le"
         Height          =   195
         Left            =   2790
         TabIndex        =   10
         Top             =   765
         Width           =   810
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
   Begin MSComctlLib.StatusBar stb_DecisionCommission 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   47
      Top             =   8610
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "03/10/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10:38"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7347
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_DecisionCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Notes: Les règles de gestion suivantes ont été retenues:
'   Un retrait d'agrément ne peut intervenir que dans le cadre d'une demande d'extension d'agrément, donc après une décision d'accord
'   Il ne peut exister qu'un retrait par demande
'   La saisie de la première commission à prendre une décision peut conduire aux décisions suivantes: Accord, Refus
'   La saisie d'une nouvelle commission est interdite dès lors qu'il existe une décision de refus ou de retrait non invalidée par recours, à moins que la décision de refus n'intervienne dans le cadre d'une extension d'agrément (donc après une décision d'accord)

Private WithEvents mfrmDemandeAdoptionEnfant As frm_DemandeAdoptionEnfant
Attribute mfrmDemandeAdoptionEnfant.VB_VarHelpID = -1

Private Enum IndexCheckBox          'Index des contrôles CheckBox
    chkFRATRIE = 0
End Enum
Private Enum IndexColumnHeader      'Index des contrôles ColumnHeader
    hdrNUMÉRO = 0
    hdrSEXE
    hdrHANDICAP
    hdrÂGE
End Enum
Private Enum IndexComboBox          'Index des contrôles ComboBox
    cboPROVENANCE = 0
End Enum
Private Enum IndexCommandButton     'Index des contrôles CommandButton
    cmdHISTORIQUE = 0
End Enum
Private Enum IndexDTPicker          'Index des contrôles DTPicker
    dtpDÉLIVRÉ = 0
    dtpRETRAIT
End Enum
Private Enum IndexLabel             'Index des contrôles Label
    lblAGRÉMENT = 0
    lblDÉLIVRÉ
    lblPROVENANCE
    lblRETRAIT
    lblPROCHAINE_DERNIÈRE
End Enum
Private Enum IndexListView          'Index des contrôles ListView
    lvwENFANTS = 0
End Enum
Private Enum IndexTextBox           'Index des contrôles TextBox
    txtDEMANDEUR = 0
End Enum
Private Enum Divers                 'Constantes diverses
    NOMBRE_ORIGINES_ETHNIQUES = 4
End Enum

Private Const mstrNOMBRE_ENFANTS As String = " enfant(s)"

Private mrstProchaineDernièreCommission As Recordset, mrstOriginesEthniques As Recordset, mrstEnfantsAccordés As Recordset
Private mbooInitialisationAvecDemandeDéfinie As Boolean

Private Sub Cmb_Deci_Click()
Const pkstrNomProcédure As String = "Cmb_Deci_Click"
Dim pbooVerrouillé As Boolean
On Error GoTo Err_Cmb_Deci_Click

'   Activer/Désactiver le motif selon la nature de la décision finale
    pbooVerrouillé = Not (Cmb_Deci.ItemData(Cmb_Deci.ListIndex) = adpREFUS Or Cmb_Deci.ItemData(Cmb_Deci.ListIndex) = adpRETRAIT)
    modCommun.DéfinirVerrouillageContrôle lbl_refus, pbooVerrouillé
    modCommun.DéfinirVerrouillageContrôle txt_motifrefus, pbooVerrouillé
    If pbooVerrouillé Then txt_motifrefus = ""

    Exit Sub
    
Err_Cmb_Deci_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


Private Sub Cmb_NumDemande_Click()
Const pkstrNomProcédure As String = "Cmb_NumDemande_Click"
Dim prstAgrément As Recordset, prstEnfantsSouhaités As Recordset
Dim pintNuméroEnfant As Integer
Dim pbooVisible As Boolean
Dim reponse As String
On Error GoTo Err_Cmb_NumDemande_Click

'   Implémenter la date de dépôt de la demande
    Implémenter "Déposée le", "True"
    
    fra_commission.Enabled = True
    fra_enfants.Enabled = True
    
'   Vérifier l'existence d'informations relatives à l'agrément
    Set prstAgrément = modDAO.rstSQLExtraction(daoEXTRACTION_013, True, gstrCONSULTATION_LIMITÉE, CLng(Cmb_NumDemande))
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
        modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
        Exit Sub
    End If
    With prstAgrément
'       Si elles existent
        If Not .RecordCount = 0 Then
'           Les afficher
            dtpAdoptions(dtpDÉLIVRÉ).Value = ![DEMA_dtmAgrementDelivrance]
            cboAdoptions(cboPROVENANCE).ListIndex = modCommun.varRecherche(cboAdoptions(cboPROVENANCE), False, ![DEMA_intAgrementProvenance])
            dtpAdoptions(dtpRETRAIT).Value = modCommun.varValeurFormatéeSiNull(![u02_dtmHistoCommissionReunion], vbDate, Null)
        End If
    End With
'   Gérer leur visibilité
    pbooVisible = Not IsNull(dtpAdoptions(dtpDÉLIVRÉ).Value)
    lblAdoptions(lblDÉLIVRÉ).Visible = pbooVisible: dtpAdoptions(dtpDÉLIVRÉ).Visible = pbooVisible
    lblAdoptions(lblPROVENANCE).Visible = pbooVisible: cboAdoptions(cboPROVENANCE).Visible = pbooVisible
    pbooVisible = Not IsNull(dtpAdoptions(dtpRETRAIT).Value)
    lblAdoptions(lblRETRAIT).Visible = pbooVisible: dtpAdoptions(dtpRETRAIT).Visible = pbooVisible
    
'   Lire la dernière/prochaine commission liée à la demande
    Set mrstProchaineDernièreCommission = modDAO.rstSQLExtraction(daoEXTRACTION_002, True, gstrCONSULTATION_LIMITÉE & gstrSÉPARATEUR_1 & gstrDEMANDE_SUPPRIMÉE, CLng(Cmb_NumDemande))
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
        modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
        Exit Sub
    End If
    
'   Positionner une liste de DTPicker à la date du jour
    modCommun.PositionnerDateJour DTP_DateDemande, DTP_DateReCom, DTP_DateDeci, DTP_PreDeci
    
'   Positionner les contrôles
    With mrstProchaineDernièreCommission
    
'       Si la date de référence pour le calcul de l'échéancier n'est pas définie
        If IsNull(![DEMA_DREF]) Then
            Err.Raise mesMESSAGE_005
        modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
        End If
        
        DTP_DateDemande.Value = modCommun.varValeurFormatéeSiNull(![DEMA_DATE], vbDate, Null)
        
        DTP_DateReCom.Value = modCommun.varValeurFormatéeSiNull(![u02_dtmHistoCommissionReunion], vbDate, Null)
        DTP_DateDeci.Value = modCommun.varValeurFormatéeSiNull(![u02_dtmHistoCommissionSignature], vbDate, Null)
'       Si aucune commission n'a encore été saisie et si la date de référence pour le calcul de l'échéancier est définie
        If IsNull(![u02_dtmHistoCommissionReunion]) And Not IsNull(![DEMA_DREF]) Then
'           Calculer la date prévue pour la commission
            DTP_PreDeci.Value = DateAdd("m", modDAO.intParamètreÉchéancier(daoRÉCEPTION_DOSSIER__DÉCISION), ![DEMA_DREF])
'       Sinon
        Else
            DTP_PreDeci.Value = modCommun.varValeurFormatéeSiNull(![u02_dtmHistoCommissionPrevu], vbDate, Null)
        End If
        
        Cmb_Avis.ListIndex = modCommun.varRecherche(Cmb_Avis, False, modCommun.varValeurFormatéeSiNull(![u02_intHistoCommissionAvis], vbInteger, AdoptionsDélibérations.adpINDÉTERMINÉ))
        Cmb_Deci.ListIndex = modCommun.varRecherche(Cmb_Deci, False, modCommun.varValeurFormatéeSiNull(![u02_intHistoCommissionDecision], vbInteger, AdoptionsDélibérations.adpINDÉTERMINÉ))
        
        txt_motifrefus = modCommun.varValeurFormatéeSiNull(![u02_txtHistoCommissionMotif], vbString)
        If Cmb_Deci.ItemData(Cmb_Deci.ListIndex) = adpREFUS Or Cmb_Deci.ItemData(Cmb_Deci.ListIndex) = adpRETRAIT Then txt_motifrefus.Enabled = True
        
        chkAdoptions(chkFRATRIE) = IIf(![u02_booHistoCommissionFratrie], vbChecked, vbUnchecked)
        
    End With
    
    modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdHISTORIQUE), IsNull(DTP_DateReCom.Value)
    
'   Si la date de commission est définie
    If Not IsNull(DTP_DateReCom.Value) Then
    
      stb_DecisionCommission.Panels(4) = "MODIFICATION"
      
'       Implémenter les origines ethniques accordées
        Implémenter "Origines ethniques", "COAC"
        If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub
        
'   Sinon
    Else
    
      stb_DecisionCommission.Panels(4).Text = "CREATION"
      
'       Si l'utilisateur décide de reprendre les souhaits exprimés lors de la demande
        If modMessages.lngRéponseUtilisateur(mesQUESTION_006) = vbYes Then
        
'           Positionner la fratrie
            chkAdoptions(chkFRATRIE) = IIf(modCommun.varValeurFormatéeSiNull(mrstProchaineDernièreCommission![DEMA_CFRAT], vbString) = "O", vbChecked, vbUnchecked)
        
'           Lire les enfants souhaités
            Set prstEnfantsSouhaités = modDAO.rstSQLExtraction(daoENFANTS_DEMANDÉS_OU_ACCORDÉS, True, gstrCONSULTATION_LIMITÉE, "FADE", CLng(Cmb_NumDemande))
            If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
                modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
                Exit Sub
            End If
            
'           Lire le numéro du dernier enfant créé
            pintNuméroEnfant = modDAO.varValeurMaximale("ENDE", "ENDE_NUM", gstrCONSULTATION_LIMITÉE)
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
                modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
                Exit Sub
            End If
            
            With prstEnfantsSouhaités
            
'               Pour chaque enfant souhaité
                While Not .EOF
            
'                   Créer l'enfant correspondant
                    Do
                        pintNuméroEnfant = pintNuméroEnfant + 1
                        modDAO.SQLAjouter daoADD_ENFANT, True, gstrCONSULTATION_LIMITÉE & gstrSÉPARATEUR_1 & gstrDOUBLON_SANS_MESSAGE, pintNuméroEnfant, ![ENDE_SEXE], ![ENDE_HAND], ![ENDE_TRAGE]
                        If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoDOUBLON) Then
                            modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
                            Exit Sub
                        End If
                    Loop While modDAO.lngDernièreErreurDAO() = daoDOUBLON
                    
'                   Créer l'enfant souhaité en commission (c'est-à-dire accordé)
                    modDAO.SQLAjouter daoADD_ENFANT_DEMANDÉ_OU_ACCORDÉ, True, gstrCONSULTATION_LIMITÉE, "COAC", CLng(Cmb_NumDemande), pintNuméroEnfant
                    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
                        modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
                        Exit Sub
                    End If
                
'            'on créera les enregistrements dans ORSO (origines ethniques) lors de la validation frm_decisioncommission
                    .MoveNext
                    
                Wend
                
            End With
            
'           Implémenter les origines ethniques souhaitées
            Implémenter "Origines ethniques", "FADE"
            If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub
                
      End If
      
    End If
        
'   Implémenter la liste des enfants accordés
    Implémenter "Enfants"
    
'   Désactiver les identifiants
    modCommun.DéfinirVerrouillageContrôle Cmb_NumDossier, True, True
    modCommun.DéfinirVerrouillageContrôle Cmb_NumDemande, True, True
  
    Exit Sub

Err_Cmb_NumDemande_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:     pintTraitement = mesINTERROMPRE
        Case mesMESSAGE_005:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
        Case Else:              pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub Cmb_NumDossier_Click()
Const pkstrNomProcédure As String = "Cmb_NumDossier_Click"
On Error GoTo Err_Cmb_NumDossier_Click

'   Afficher l'identité du demandeur
    modAdoptions.AfficherIdentitéDemandeur Cmb_NumDossier, txtAdoptions(txtDEMANDEUR), gstrCONSULTATION_LIMITÉE, cmd_Valider
        
'   Implémenter la date de dépôt de la demande
    Implémenter "Déposée le", "False"
    
'   Si la demande n'est pas définie à l'initialisation de la fenêtre, lire les numéros des demandes liées au dossier
    If Not mbooInitialisationAvecDemandeDéfinie Then modAdoptions.ImplémenterListe Cmb_NumDemande, "Demandes", gstrCONSULTATION_LIMITÉE, cmd_Valider, Cmb_NumDossier
    
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


Private Sub cmd_CreerIndividu_Click()
Const pkstrNomProcédure As String = "cmd_CreerIndividu_Click"
On Error GoTo Err_cmd_CreerIndividu_Click

'   S'il existe déjà un enfant au moins et que la fratrie n'est pas autorisée
    If lvwAdoptions(lvwENFANTS).ListItems.Count > 0 And chkAdoptions(chkFRATRIE) = vbUnchecked Then
        Err.Raise mesFRATRIE_REFUSÉE
        Exit Sub
'   Sinon
    Else
'       Ouvrir la fenêtre permettant le traitement des enfants
        With mfrmDemandeAdoptionEnfant
            .Initialiser comCRÉATION, False, Cmb_NumDemande
            .Show vbModal
        End With
        
    End If
    
    Exit Sub
    
Err_cmd_CreerIndividu_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:         pintTraitement = mesINTERROMPRE
        Case mesFRATRIE_REFUSÉE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, chkAdoptions(chkFRATRIE), "Création")
        Case Else:                  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        

End Sub

Private Sub cmd_ModifierEnfant_Click()
Const pkstrNomProcédure As String = "cmd_ModifierEnfant_Click"
Dim pitmEnfant As ListItem
On Error GoTo Err_cmd_ModifierEnfant_Click

'   Référencer l'enfant sélectionné
    Set pitmEnfant = lvwAdoptions(lvwENFANTS).SelectedItem
    
'   S'il existe
    If Not pitmEnfant Is Nothing Then
'       Ouvrir la fenêtre permettant le traitement des enfants
        With mfrmDemandeAdoptionEnfant
            .Initialiser comMODIFICATION, False, Cmb_NumDemande, pitmEnfant.Text
            .Show vbModal
        End With
    End If
    
    Exit Sub
    
Err_cmd_ModifierEnfant_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_Quitter_Click()
Const pkstrNomProcédure As String = "cmd_Quitter_Click"

    Select Case stb_DecisionCommission.Panels(4)
    
'       L'utilisateur renonce à créer une première commission
        Case "CREATION"
        
'           Ouvrir une transaction
            modDAO.GérerTransaction daoOUVERTURE
'           Supprimer les enfants accordés pour la demande
            SupprimerEnfantsAccordés "La fermeture de la fenêtre est refusée."
            If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then modDAO.GérerTransaction daoANNULATION: Exit Sub
'           Valider la transaction
            modDAO.GérerTransaction daoVALIDATION
        
'       L'utilisateur renonce à mettre à jour la commission
        Case "MODIFICATION"
        
'           Refuser la fermeture de la fenêtre s'il n'a pas été saisi de profil d'enfant
            If booUnloadRefusé(False) Then Exit Sub
        
    End Select
    
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

Private Sub cmd_SupprimerEnfant_Click()
Const pkstrNomProcédure As String = "cmd_SupprimerEnfant_Click"
Dim pitmEnfant As ListItem
On Error GoTo Err_cmd_SupprimerEnfant_Click

'   Référencer l'enfant sélectionné
    Set pitmEnfant = lvwAdoptions(lvwENFANTS).SelectedItem
    
'   S'il existe
    If Not pitmEnfant Is Nothing Then
'       Ouvrir la fenêtre permettant le traitement des enfants
        With mfrmDemandeAdoptionEnfant
            .Initialiser comSUPPRESSION, False, Cmb_NumDemande, pitmEnfant.Text
            .Show vbModal
        End With
    End If
    
    Exit Sub
    
Err_cmd_SupprimerEnfant_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_Valider_Click()
Const pkstrNomProcédure As String = "cmd_Valider_Click"
Dim pvarRecoursAdministratif As Variant, pvarRecoursContentieux As Variant
Dim prstRecoursAdministratif As Recordset, prstRecoursContentieux As Recordset
Dim pitmEnfant As ListItem
Dim plngTraitement As CommunTraitements
Dim pintItération As Integer
Dim pbooSuppressionDemandée As Boolean
On Error GoTo Err_cmd_Valider_Click
  
'   Forcer si nécessaire l'actualisation de la valeur d'un contrôle DTPicker
    modCommun.GérerActualisationDTPicker Me, cmd_Valider
    
'   Refuser la validation si le numéro de la demande n'est pas sélectionné
    If Cmb_NumDemande = "" Then
        Err.Raise mesSÉLECTION_MANQUANTE
        Exit Sub
    End If
    
'   Si la date prévue pour la commission a été désactivée
    If IsNull(DTP_PreDeci.Value) Then
'       La recalculer
        DTP_PreDeci.Value = DateAdd("m", modDAO.intParamètreÉchéancier(daoRÉCEPTION_DOSSIER__DÉCISION), mrstProchaineDernièreCommission![DEMA_DREF])
    End If
    
'   Ouvrir une transaction
    modDAO.GérerTransaction daoOUVERTURE
    
'   Si la date de la commission est définie
    If Not IsNull(DTP_DateReCom.Value) Then
    
'       S'il n'a pas été saisi de profil d'enfant
        If booUnloadRefusé(True) Then
            modDAO.GérerTransaction daoANNULATION: Exit Sub
        End If
    
'       Refuser la validation si la fratrie n'est pas autorisée et plus d'un enfant saisi
        If chkAdoptions(chkFRATRIE) = vbUnchecked And lvwAdoptions(lvwENFANTS).ListItems.Count > 1 Then
            Err.Raise mesFRATRIE_REFUSÉE
            modDAO.GérerTransaction daoANNULATION: Exit Sub
        End If
    
        With mrstProchaineDernièreCommission
        
'           S'il s'agit de la saisie d'une première commission
            If IsNull(![u02_dtmHistoCommissionReunion]) Then
                plngTraitement = comCRÉATION
'           Sinon, si la date de la commission a été modifiée
            ElseIf Not DTP_DateReCom.Value = ![u02_dtmHistoCommissionReunion] Then
'               Proposer à l'utilisateur l'enregistrement d'une nouvelle commission
                Select Case modMessages.lngRéponseUtilisateur(mesQUESTION_005, CStr(![u02_dtmHistoCommissionReunion]))
                    Case vbYes:     plngTraitement = comCRÉATION
                    Case vbNo:      plngTraitement = comMODIFICATION
                    Case vbCancel:  modDAO.GérerTransaction daoANNULATION: Exit Sub
                End Select
'           Sinon
            Else
                plngTraitement = comMODIFICATION
            End If
            
'           Vérifier la cohérence de la décision
            If modAdoptions.booSaisieIncohérenteDécision(CLng(Cmb_NumDemande), gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_2 & CStr(vbInformation), plngTraitement, ![u02_dtmHistoCommissionReunion], AdoptionsProvenances.adpDEMANDE_INITIALE, DTP_DateReCom, Cmb_Deci, DTP_DateDeci, pdtpDécisionPrévue:=DTP_PreDeci, pcmdÀDésactiver:=cmd_Valider) Then modDAO.GérerTransaction daoANNULATION: Exit Sub

'           S'il s'agit d'enregistrer une nouvelle commission
            If plngTraitement = comCRÉATION Then
            
'               Créer l'historique
                modDAO.SQLAjouter daoAJOUT_001, True, gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_1 & "Création refusée: Une commission existe déjà pour cette date.", CLng(Cmb_NumDemande), DTP_DateReCom.Value, DTP_DateDeci.Value, DTP_PreDeci.Value, Cmb_Avis.ItemData(Cmb_Avis.ListIndex), Cmb_Deci.ItemData(Cmb_Deci.ListIndex), IIf(txt_motifrefus = "", Null, txt_motifrefus), (chkAdoptions(chkFRATRIE) = vbChecked)
                If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                
'           Sinon
            Else
            
'               Vérifier l'existence d'un recours administratif
                Set prstRecoursAdministratif = modDAO.rstSQLExtraction(daoRECOURS_ADMINISTRATIF, True, gstrVALIDATION_ANNULÉE, dbOpenSnapshot, CLng(Cmb_NumDemande), ![u02_dtmHistoCommissionReunion])
                If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                
'               S'il existe un recours administratif
                If Not modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT Then
'                   Le copier dans un tableau
                    With prstRecoursAdministratif
                        .MoveFirst
                        pvarRecoursAdministratif = .GetRows(.RecordCount)
                    End With
'                   Le supprimer
                    modDAO.SQLSupprimer daoDEL_HISTORIQUE_COMMISSION_OU_RECOURS, True, gstrVALIDATION_ANNULÉE, "READ", CLng(Cmb_NumDemande), ![u02_dtmHistoCommissionReunion]
                    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                End If
                    
'               Vérifier l'existence de recours contentieux
                Set prstRecoursContentieux = modDAO.rstSQLExtraction(daoRECOURS_CONTENTIEUX, True, gstrVALIDATION_ANNULÉE, dbOpenSnapshot, CLng(Cmb_NumDemande), ![u02_dtmHistoCommissionReunion])
                If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then modDAO.GérerTransaction daoANNULATION: Exit Sub

'               S'il existe un ou plusieurs recours contentieux
                If Not modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT Then
'                   Les copier dans un tableau
                    With prstRecoursContentieux
                        .MoveFirst
                        pvarRecoursContentieux = .GetRows(.RecordCount)
                    End With
'                   Les supprimer
                    modDAO.SQLSupprimer daoDEL_HISTORIQUE_COMMISSION_OU_RECOURS, True, gstrVALIDATION_ANNULÉE, "RECO", CLng(Cmb_NumDemande), ![u02_dtmHistoCommissionReunion]
                    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                End If
                    
'               Mettre à jour l'historique
                modDAO.SQLMettreÀJour daoMISE_À_JOUR_001, True, gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_1 & "Mise à jour refusée: Une commission existe déjà pour cette date.", CLng(Cmb_NumDemande), ![u02_dtmHistoCommissionReunion], DTP_DateReCom.Value, DTP_DateDeci.Value, DTP_PreDeci.Value, Cmb_Avis.ItemData(Cmb_Avis.ListIndex), Cmb_Deci.ItemData(Cmb_Deci.ListIndex), IIf(txt_motifrefus = "", Null, txt_motifrefus), (chkAdoptions(chkFRATRIE) = vbChecked)
                If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                
'               S'il existait un recours administratif
                If Not IsEmpty(pvarRecoursAdministratif) Then
'                   Le recréer
                    modDAO.SQLAjouter daoADD_RECOURS_ADMINISTRATIF, True, gstrVALIDATION_ANNULÉE, pvarRecoursAdministratif(0, 0), pvarRecoursAdministratif(1, 0), pvarRecoursAdministratif(2, 0), pvarRecoursAdministratif(3, 0), pvarRecoursAdministratif(4, 0), pvarRecoursAdministratif(5, 0), pvarRecoursAdministratif(6, 0), pvarRecoursAdministratif(7, 0), pvarRecoursAdministratif(8, 0), pvarRecoursAdministratif(9, 0), pvarRecoursAdministratif(10, 0), pvarRecoursAdministratif(11, 0), pvarRecoursAdministratif(12, 0), pvarRecoursAdministratif(13, 0), pvarRecoursAdministratif(14, 0), pvarRecoursAdministratif(15, 0), DTP_DateReCom.Value
                    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                End If
                    
'               S'il existait un ou plusieurs recours contentieux
                If Not IsEmpty(pvarRecoursContentieux) Then
'                   Les recréer
                    For pintItération = 0 To UBound(pvarRecoursContentieux, 2)
                        modDAO.SQLAjouter daoAJOUT_002, True, gstrVALIDATION_ANNULÉE, pvarRecoursContentieux(0, pintItération), pvarRecoursContentieux(1, pintItération), pvarRecoursContentieux(2, pintItération), pvarRecoursContentieux(3, pintItération), pvarRecoursContentieux(4, pintItération), pvarRecoursContentieux(5, pintItération), pvarRecoursContentieux(6, pintItération), pvarRecoursContentieux(7, pintItération), DTP_DateReCom.Value
                        If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                    Next pintItération
                End If
                    
            End If
            
        End With

'       Pour chaque enfant accordé
        For Each pitmEnfant In lvwAdoptions(lvwENFANTS).ListItems
        
'           Supprimer ses orgines ethniques
            modDAO.SQLSupprimer daoDEL_ORIGINES_OU_ENFANT_OU_DEMANDÉ_OU_ACCORDÉ, True, gstrVALIDATION_ANNULÉE, "ORSO", CInt(pitmEnfant.Text)
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
            
'           Pour chaque origine ethnique
            For pintItération = 0 To CInt(adpNOMBRE_ORIGINES_ETHNIQUES) - 1
'               Si elle est sélectionnée
                If Not Cmb_OrigEt(pintItération).ItemData(Cmb_OrigEt(pintItération).ListIndex) = adpAUCUNE_SÉLECTION Then
'                   L'associer à l'enfant
                    modDAO.SQLAjouter daoAJOUT_003, True, gstrVALIDATION_ANNULÉE, Cmb_OrigEt(pintItération).ItemData(Cmb_OrigEt(pintItération).ListIndex), CInt(pitmEnfant.Text)
                    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                End If
            Next pintItération
            
        Next pitmEnfant

'   Sinon
    Else
    
        With mrstProchaineDernièreCommission
        
'           Si aucune commission n'est enregistrée
            If IsNull(![u02_dtmHistoCommissionReunion]) Then
                Err.Raise mesSAISIE_MANQUANTE
                modDAO.GérerTransaction daoANNULATION: Exit Sub
            
'           Sinon
            Else
            
'               Vérifier la cohérence de la suppression
                If modAdoptions.booSaisieIncohérenteDécision(CLng(Cmb_NumDemande), gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_2 & CStr(vbInformation), comSUPPRESSION, ![u02_dtmHistoCommissionReunion], AdoptionsProvenances.adpDEMANDE_INITIALE, DTP_DateReCom, Cmb_Deci, DTP_DateDeci, pdtpDécisionPrévue:=DTP_PreDeci, pcmdÀDésactiver:=cmd_Valider) Then modDAO.GérerTransaction daoANNULATION: Exit Sub

'               Demander confirmation de la suppression de l'inscription du dossier à la commission
                Select Case modMessages.lngRéponseUtilisateur(mesSUPPRIMER_INSCRIPTION_COMMISSION, ![u02_dtmHistoCommissionReunion])
'                   L'utilisateur confirme la suppression
                    Case vbYes
                        pbooSuppressionDemandée = True
'                       Supprimer la commission
                        modDAO.SQLSupprimer daoDEL_HISTORIQUE_COMMISSION_OU_RECOURS, True, gstrVALIDATION_ANNULÉE, "u02_HistoCommissions", CLng(Cmb_NumDemande), ![u02_dtmHistoCommissionReunion]
                        If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
'                       Actualiser la dernière/prochaine commission liée à la demande
                        .Requery
'                       S'il n'en existe plus
                        If IsNull(![u02_dtmHistoCommissionReunion]) Then
'                           Supprimer les enfants accordés pour la demande
                            SupprimerEnfantsAccordés gstrVALIDATION_ANNULÉE
                            If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
                                modDAO.GérerTransaction daoANNULATION
                                .Requery
                                Exit Sub
                            End If
                        End If
'                   L'utilisateur renonce à la suppression
                    Case vbNo
                        modDAO.GérerTransaction daoANNULATION: Exit Sub
                End Select
                
            End If
        
        End With
        
    End If

'   Actualiser les données relatives à l'agrément
    modAdoptions.ActualiserDonnéesAgrément CLng(Cmb_NumDemande), gstrVALIDATION_ANNULÉE
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
        modDAO.GérerTransaction daoANNULATION
'       Si la supression a été demandée, actualiser la dernière/prochaine commission liée à la demande
'           Note: Permet de gérer le cas d'une suppression quand la commission était la seule liée à la demande
        If pbooSuppressionDemandée Then mrstProchaineDernièreCommission.Requery
        Exit Sub
    End If

'   Valider la transaction
    modDAO.GérerTransaction daoVALIDATION
        
    Unload Me
    
    Exit Sub
  
Err_cmd_Valider_Click:
Dim pintTraitement As MessagesActions
    Select Case Err
        Case vbObjectError:             pintTraitement = mesINTERROMPRE
        Case mesSÉLECTION_MANQUANTE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Cmb_NumDemande, "du numéro de demande")
        Case mesSAISIE_MANQUANTE:       pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, DTP_DateReCom, "de la date de la commission")
        Case mesFRATRIE_REFUSÉE:        pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, chkAdoptions(chkFRATRIE), "Validation")
        Case Else:                      pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub

Private Sub cmdAdoptions_Click(Index As Integer)
Const pkstrNomProcédure As String = "cmdAdoptions_Click"
Dim prstNombrePlacements As Recordset
On Error GoTo Err_cmdAdoptions_Click
    
    Select Case Index
    
'       Visualiser l'historique des commissions
        Case cmdHISTORIQUE
        
            frmHistoriqueCommissions.Initialiser CLng(Cmb_NumDemande)
            frmHistoriqueCommissions.Show vbModal
            
    End Select
        
    Exit Sub
    
Err_cmdAdoptions_Click:
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

    Set mrstProchaineDernièreCommission = Nothing: Set mrstOriginesEthniques = Nothing: Set mrstEnfantsAccordés = Nothing
    Set mfrmDemandeAdoptionEnfant = Nothing
    
    Exit Sub
    
Err_Form_Unload:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub mfrmDemandeAdoptionEnfant_Validation(plngTraitement As CommunTraitements, pintNuméroEnfant As Integer)
Const pkstrNomProcédure As String = "mfrmDemandeAdoptionEnfant_Validation"
Dim pitmEnfant As ListItem
On Error GoTo Err_mfrmDemandeAdoptionEnfant_Validation

'   Implémenter la liste des enfants accordés
    Implémenter "Enfants"
    
'   S'il ne s'agit pas d'une suppression
    If Not (plngTraitement = comSUPPRESSION Or pintNuméroEnfant = CInt(daoENREGISTREMENT_SUPPRIMÉ)) Then
'       Référencer l'enfant sur lequel a porté l'opération
        Set pitmEnfant = lvwAdoptions(lvwENFANTS).ListItems("*" & CStr(pintNuméroEnfant))
        pitmEnfant.Selected = True
    End If
    
    Exit Sub
    
Err_mfrmDemandeAdoptionEnfant_Validation:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Retourner un booléen précisant si le déchargement de la feuille est refusé ou non
'   pbooValidation: booléen précisant si le déchargement doit intervenir dans le cadre d'une validation

Private Function booUnloadRefusé(pbooValidation As Boolean) As Boolean
Const pkstrNomProcédure As String = "booUnloadRefusé"
Dim pctlFocus As Control
Dim pintItération As Integer
Dim pbooOrigineEthniqueSaisie As Boolean
On Error GoTo Err_booUnloadRefusé

'   Si aucun enfant n'a été saisi
    If lvwAdoptions(lvwENFANTS).ListItems.Count = 0 Then
        Set pctlFocus = cmd_CreerIndividu
        booUnloadRefusé = True
'   Sinon
    Else
        Set pctlFocus = Cmb_OrigEt(0)
'       S'il s'agit de valider la saisie
        If pbooValidation Then
'           Vérifier la saisie d'une origine ethnique
            For pintItération = 0 To CInt(adpNOMBRE_ORIGINES_ETHNIQUES) - 1
                If Not Cmb_OrigEt(pintItération).ItemData(Cmb_OrigEt(pintItération).ListIndex) = adpAUCUNE_SÉLECTION Then pbooOrigineEthniqueSaisie = True: Exit For
            Next pintItération
            booUnloadRefusé = Not pbooOrigineEthniqueSaisie
'       Sinon
        Else
'           Implémenter les origines ethniques accordées
            Implémenter "Origines ethniques", "COAC"
            booUnloadRefusé = (mrstOriginesEthniques.RecordCount = 0)
        End If
    End If

'   Refuser la validation s'il n'a pas été saisi de profil d'enfant
    If booUnloadRefusé Then Err.Raise mesMESSAGE_006

    Exit Function
    
Err_booUnloadRefusé:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:     pintTraitement = mesINTERROMPRE
        Case mesMESSAGE_006:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, IIf(pbooValidation, "Validation", "Annulation"))
        Case Else:              pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Function


'Implémenter les objets
'   pstrListeObjets:        liste des objets à implémenter au format '<objet#1>[|<objet#2>[...|<objet#n]]'
'   pstrListeParamètres:    liste des paramètres au format '<paramètre#1 de l'objet#1[¤paramètre#2 de l'objet#1[...¤paramètre#n de l'objet#1]]>[|<paramètre#1 de l'objet#2[¤paramètre#2 de l'objet#2[...¤paramètre#n de l'objet#2]]>[...|<paramètre#1 de l'objet#n[¤paramètre#2 de l'objet#n[...¤paramètre#n de l'objet#n]]>]]'

Private Sub Implémenter(pstrListeObjets As String, Optional pstrListeParamètres As String = "")
Const pkstrNomProcédure As String = "Implémenter"
Const pkintNom As Integer = 0, pkintIndex As Integer = 1
Dim pitmEnfant As ListItem
Dim pstrListesObjets() As String, pstrListesParamètres() As String, pstrParamètres() As String
Dim pintItérationObjet As Integer, pintItérationOrigine As Integer
On Error GoTo Err_Implémenter
    
    pstrListesObjets() = Split(pstrListeObjets, ","): pstrListesParamètres() = Split(pstrListeParamètres, gstrSÉPARATEUR_1)
    
'   Pour chaque objet dont on souhaite l'implémentation
    For pintItérationObjet = 0 To UBound(pstrListesObjets)
    
        Select Case pstrListesObjets(pintItérationObjet)
        
'           Date de dépôt de la demande
            Case "Déposée le"
        
'               Positionner la visibilité
                pstrParamètres() = Split(pstrListesParamètres(pintItérationObjet), gstrSÉPARATEUR_2)
                DTP_DateDemande.Visible = CBool(pstrParamètres(0)): Lbl_DateDemande.Visible = CBool(pstrParamètres(0))
        
'           Origines ethniques (souhaitées ou accordées)
            Case "Origines ethniques"
        
                pstrParamètres() = Split(pstrListesParamètres(pintItérationObjet), gstrSÉPARATEUR_2)
        
'               Si les origines ethniques n'ont pas encore été lues
                If mrstOriginesEthniques Is Nothing Then
'                   Les lire
                    Set mrstOriginesEthniques = modDAO.rstSQLExtraction(daoEXTRACTION_020, True, gstrCONSULTATION_LIMITÉE, pstrParamètres(0), CLng(Cmb_NumDemande))
'               Sinon
                Else
'                   Rafraîchir le jeu d'enregistrements
                    modDAO.ActualiserRecordset mrstOriginesEthniques, True, gstrCONSULTATION_LIMITÉE
                End If
                If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
                    modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
                    Exit Sub
                End If
                
                For pintItérationOrigine = 0 To NOMBRE_ORIGINES_ETHNIQUES - 1
                    Cmb_OrigEt(pintItérationOrigine).ListIndex = modCommun.varRecherche(Cmb_OrigEt(pintItérationOrigine), False, adpAUCUNE_SÉLECTION)
                Next pintItérationOrigine

                With mrstOriginesEthniques
'                   Pour chacune
                    While Not .EOF
                        Cmb_OrigEt(.AbsolutePosition).ListIndex = modCommun.varRecherche(Cmb_OrigEt(.AbsolutePosition), False, ![ORET_CODE])
                        .MoveNext
                    Wend
                End With
                
                If pstrParamètres(0) = "FADE" Then Set mrstOriginesEthniques = Nothing
                
'           Enfants accordés
            Case "Enfants"
        
                lvwAdoptions(lvwENFANTS).ListItems.Clear
                
'               Si les enfants accordés n'ont pas encore été lus
                If mrstEnfantsAccordés Is Nothing Then
'                   Les lire
                    Set mrstEnfantsAccordés = modDAO.rstSQLExtraction(daoENFANTS_DEMANDÉS_OU_ACCORDÉS, True, gstrCONSULTATION_LIMITÉE, "COAC", CLng(Cmb_NumDemande))
'               Sinon
                Else
'                   Rafraîchir le jeu d'enregistrements
                    modDAO.ActualiserRecordset mrstEnfantsAccordés, True, gstrCONSULTATION_LIMITÉE
                End If
                If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
                    modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
                    Exit Sub
                End If
                
                With mrstEnfantsAccordés
'                   Pour chacun
                    While Not .EOF
'                       Créer un item sans image
                        Set pitmEnfant = lvwAdoptions(lvwENFANTS).ListItems.Add(, "*" & CStr(![ENDE_NUM]), CStr(![ENDE_NUM]))
                        modCommun.ImplémenterSubItems pitmEnfant, modCommun.varValeurFormatéeSiNull(![ENDE_SEXE], vbString), modCommun.varValeurFormatéeSiNull(![ENDE_HAND], vbString), modCommun.varValeurFormatéeSiNull(![ENDE_TRAGE], vbString)
                        .MoveNext
                    Wend
                    Txt_NombreEnfants = CStr(.RecordCount) & mstrNOMBRE_ENFANTS
               End With
        
        End Select
        
    Next pintItérationObjet
    
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
Dim pstrMessages(2) As String, pstrNuméroDossier As String
Dim pintItération As Integer
On Error GoTo Err_Initialiser
    
    mbooInitialisationAvecDemandeDéfinie = Not IsMissing(pvarNuméroDemande)
    
'   Centrer la feuille
    modAdoptions.CentrerFeuille Me
        
'   Implémenter les images
    cmd_Quitter.Picture = Menu.picImage(adp32x32, adpANNULER)
    cmd_Valider.Picture = Menu.picImage(adp32x32, adpVALIDER)
    cmdAdoptions(cmdHISTORIQUE).Picture = Menu.picImage(adpTEXTE_IMAGE, adpTEXTE_HISTORIQUE)
    cmd_CreerIndividu.Picture = Menu.picImage(adp16x16, adpCRÉER)
    cmd_ModifierEnfant.Picture = Menu.picImage(adp16x16, adpMODIFIER)
    cmd_SupprimerEnfant.Picture = Menu.picImage(adp16x16, adpSUPPRIMER)
        
'   Désactiver certains contrôles
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtDEMANDEUR), True
    modCommun.DéfinirVerrouillageContrôle DTP_DateDemande, True
    modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpDÉLIVRÉ), True
    modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpRETRAIT), True
    modCommun.DéfinirVerrouillageContrôle cboAdoptions(cboPROVENANCE), True
    
'   Vérifier le paramétrage des recours contentieux
    With gclsRecoursContentieux
        If .booParamétrageIncomplet Then
            pstrMessages(0) = .strTribunalAdministratif: pstrMessages(1) = .strCourAdministrativeAppel: pstrMessages(2) = .strConseilÉtat
            Err.Raise mesPARAMÉTRAGE_CONTENTIEUX_INCOMPLET
            modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
            Exit Sub
        End If
    End With

    Set mfrmDemandeAdoptionEnfant = New frm_DemandeAdoptionEnfant
    
'   Implémenter les origines ethniques
    For pintItération = 0 To CInt(adpNOMBRE_ORIGINES_ETHNIQUES) - 1
        modAdoptions.ImplémenterListe Cmb_OrigEt(pintItération), "Origines ethniques", gstrCONSULTATION_LIMITÉE, cmd_Valider
        Cmb_OrigEt(pintItération).ListIndex = modCommun.varRecherche(Cmb_OrigEt(pintItération), False, adpAUCUNE_SÉLECTION)
    Next pintItération
'   Implémenter la provenance
    modAdoptions.ImplémenterListe cboAdoptions(cboPROVENANCE), "Provenances", "", Nothing
    cboAdoptions(cboPROVENANCE).ListIndex = modCommun.varRecherche(cboAdoptions(cboPROVENANCE), False, adpINDÉTERMINÉE)
'   Implémenter l'avis
    modAdoptions.ImplémenterListe Cmb_Avis, "Délibérations", "", Nothing, True
    Cmb_Avis.ListIndex = modCommun.varRecherche(Cmb_Avis, False, AdoptionsDélibérations.adpINDÉTERMINÉ)
'   Implémenter la décision finale
    modAdoptions.ImplémenterListe Cmb_Deci, "Délibérations", "", Nothing, True
    Cmb_Deci.ListIndex = modCommun.varRecherche(Cmb_Deci, False, AdoptionsDélibérations.adpINDÉTERMINÉ)
    
    dtpAdoptions(dtpDÉLIVRÉ) = "": dtpAdoptions(dtpRETRAIT) = ""
    
'   Implémenter la date de dépôt de la demande
    Implémenter "Déposée le", "False"
    
'   Si la demande est définie
    If Not IsMissing(pvarNuméroDemande) Then
        pstrNuméroDossier = modAdoptions.txtNuméroDossier(CLng(pvarNuméroDemande), cmd_Valider)
        If pstrNuméroDossier = "" Then Exit Sub
        modCommun.ImplémenterListe Cmb_NumDossier, pstrNuméroDossier
        modCommun.ImplémenterListe Cmb_NumDemande, CStr(pvarNuméroDemande)
        Cmb_NumDossier.ListIndex = modCommun.varRecherche(Cmb_NumDossier, True, pstrNuméroDossier)
        Cmb_NumDemande.ListIndex = modCommun.varRecherche(Cmb_NumDemande, True, CStr(pvarNuméroDemande))
'   Sinon
    Else
        modAdoptions.ImplémenterListe Cmb_NumDossier, "Dossiers", gstrCONSULTATION_LIMITÉE, cmd_Valider
    End If
    
    Exit Sub
    
Err_Initialiser:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                         pintTraitement = mesINTERROMPRE
        Case mesPARAMÉTRAGE_CONTENTIEUX_INCOMPLET:  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, pstrMessages(0), pstrMessages(1), pstrMessages(2))
        Case Else:                                  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Supprimer les enfants accordés
'   Les valeurs prises par modDAO.lngDernièreErreurDAO() sont:  (*) daoENREGISTREMENT_VERROUILLÉ:   Des données sont verrouillées par une transaction
'                                                               (*) daoAUCUN_ENREGISTREMENT:        Aucun enfant n'est accordé pour la demande
'                                                               (*) daoAUCUNE:                      L'(Les )enfant(s) a(ont) été supprimé(s)
'       (*) la valeur est retournée et peut donc être interceptée par la procédure appelante
'   pstrMessageSiVerrouillage:  message à afficher si les données sont verrouillées par une transaction

Private Sub SupprimerEnfantsAccordés(pstrMessageSiVerrouillage As String)
Const pkstrNomProcédure As String = "SupprimerEnfantsAccordés"
Dim prstEnfantsAccordés As Recordset
On Error GoTo Err_SupprimerEnfantsAccordés
    
'   Lire les enfants accordés
    Set prstEnfantsAccordés = modDAO.rstSQLExtraction(daoEXTRACTION_015, True, pstrMessageSiVerrouillage, CLng(Cmb_NumDemande))
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then Exit Sub

    With prstEnfantsAccordés
'       Pour chacun
        While Not .EOF
'           Le supprimer
            modDAO.SQLSupprimer daoDEL_ORIGINES_OU_ENFANT_OU_DEMANDÉ_OU_ACCORDÉ, True, pstrMessageSiVerrouillage, "ORSO", ![ENDE_NUM]
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then Exit Sub
            modDAO.SQLSupprimer daoDEL_ORIGINES_OU_ENFANT_OU_DEMANDÉ_OU_ACCORDÉ, True, pstrMessageSiVerrouillage, "COAC", ![ENDE_NUM]
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then Exit Sub
            modDAO.SQLSupprimer daoDEL_ORIGINES_OU_ENFANT_OU_DEMANDÉ_OU_ACCORDÉ, True, pstrMessageSiVerrouillage, "ENDE", ![ENDE_NUM]
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then Exit Sub
            .MoveNext
        Wend
    End With
    
    Exit Sub
    
Err_SupprimerEnfantsAccordés:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

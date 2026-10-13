VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_DossierGestion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion des dossiers"
   ClientHeight    =   7185
   ClientLeft      =   405
   ClientTop       =   1815
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_DossierGestion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra_Dossiers 
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
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton Cmd_CreerDossier 
         Height          =   375
         Left            =   5160
         Picture         =   "frm_DossierGestion.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Créer un nouveau dossier"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_ModifierDossier 
         Height          =   375
         Left            =   5760
         Picture         =   "frm_DossierGestion.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Modifier les informations sur le dossier sélectionné"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_SupprimerDossier 
         Height          =   375
         Left            =   6360
         Picture         =   "frm_DossierGestion.frx":0ACE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Supprimer le dossier sélectionné"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Listedossiers 
         Caption         =   "Liste des dossiers ..."
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         ToolTipText     =   "Permet de sélectionner un dossier en visualisant la liste des dossiers existants"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8760
         Picture         =   "frm_DossierGestion.frx":0BD0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Actualiser 
         Default         =   -1  'True
         Height          =   375
         Left            =   7320
         Picture         =   "frm_DossierGestion.frx":0CD2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Mise à jour de l'affichage avec les informations sur le dossier sélectionné"
         Top             =   150
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox Cmb_NumDossier 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Choisissez un numéro de dossier"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lbl_Numéro 
         Caption         =   "Numéro dossier"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar stb_GestionDossier 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   6930
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "08/07/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16:41"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8678
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SST_Dossier 
      Height          =   6135
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "Permet de visionner les caractéristiques d'un dossier et ses demandes"
      Top             =   720
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dossier"
      TabPicture(0)   =   "frm_DossierGestion.frx":0DD4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Dat_Individus"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra_Individus"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_Caracteristiques"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Fra_Situation"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Demandes"
      TabPicture(1)   =   "frm_DossierGestion.frx":0DF0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_Outils"
      Tab(1).Control(1)=   "Fra_Demandes"
      Tab(1).Control(2)=   "stb_Individu"
      Tab(1).ControlCount=   3
      Begin VB.Frame Fra_Situation 
         Caption         =   "Situation"
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
         TabIndex        =   51
         Top             =   2400
         Width           =   9255
         Begin VB.TextBox Txt_MoisMoyenne 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2880
            TabIndex        =   56
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox Txt_AnneeMoyenne 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   55
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox Txt_RegimeMatrimonial 
            Height          =   285
            Left            =   5640
            TabIndex        =   54
            Top             =   240
            Width           =   3015
         End
         Begin VB.TextBox Txt_LieuMariage 
            Height          =   285
            Left            =   5640
            TabIndex        =   53
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox Txt_SituationFamiliale 
            Height          =   285
            Left            =   1560
            TabIndex        =   52
            Top             =   240
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker DTP_DateMariage 
            Height          =   255
            Left            =   1560
            TabIndex        =   57
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yyy"
            Format          =   58982403
            CurrentDate     =   36283
         End
         Begin VB.Label lbl_age3 
            Caption         =   "mois."
            Height          =   255
            Left            =   3360
            TabIndex        =   64
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lbl_age2 
            Caption         =   "ans et"
            Height          =   255
            Left            =   2280
            TabIndex        =   63
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lbl_age 
            Caption         =   "Age moyen du couple"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lbl_regimematrimonial 
            Caption         =   "Régime matrimonial"
            Height          =   255
            Left            =   3960
            TabIndex        =   61
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lbl_lieumariage 
            Caption         =   "Lieu de mariage"
            Height          =   255
            Left            =   3960
            TabIndex        =   60
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lbl_datemariage 
            Caption         =   "Date de mariage"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lbl_situationfamiliale 
            Caption         =   "Situation familiale"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fra_Caracteristiques 
         Caption         =   "Caractéristiques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   9255
         Begin VB.TextBox Txt_Departement 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7680
            TabIndex        =   40
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Txt_Circonscription 
            Height          =   285
            Left            =   6120
            TabIndex        =   39
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox Txt_Canton 
            Height          =   285
            Left            =   6120
            TabIndex        =   38
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox Txt_Commune 
            Height          =   285
            Left            =   6120
            TabIndex        =   37
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox Txt_reinstruction 
            Height          =   285
            Left            =   6120
            TabIndex        =   36
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox Txt_Ville 
            Height          =   285
            Left            =   2640
            TabIndex        =   35
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox Txt_Tel 
            Height          =   285
            Left            =   1080
            MaxLength       =   14
            TabIndex        =   34
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox Txt_CP 
            Height          =   285
            Left            =   1080
            TabIndex        =   33
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox Txt_adr2 
            Height          =   285
            Left            =   1080
            MaxLength       =   30
            TabIndex        =   32
            Top             =   600
            Width           =   3615
         End
         Begin VB.TextBox Txt_adr1 
            Height          =   285
            Left            =   1080
            MaxLength       =   30
            TabIndex        =   31
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label lbl_CodeDepartement 
            Caption         =   "Département"
            Enabled         =   0   'False
            Height          =   255
            Left            =   6600
            TabIndex        =   50
            Top             =   240
            Width           =   1095
         End
         Begin VB.Line Lin_Separateur 
            BorderColor     =   &H80000016&
            X1              =   4920
            X2              =   4920
            Y1              =   120
            Y2              =   1800
         End
         Begin VB.Label lbl_circonscription 
            Caption         =   "Circonscription"
            Height          =   255
            Left            =   5040
            TabIndex        =   49
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lbl_canton 
            Caption         =   "Canton"
            Height          =   255
            Left            =   5040
            TabIndex        =   48
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lbl_commune 
            Caption         =   "Commune"
            Height          =   255
            Left            =   5040
            TabIndex        =   47
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lbl_reinstruction 
            Caption         =   "Réinstruction"
            Height          =   255
            Left            =   5040
            TabIndex        =   46
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lbl_tel 
            Caption         =   "Téléphone"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lbl_ville 
            Caption         =   "Ville"
            Height          =   255
            Left            =   2280
            TabIndex        =   44
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lbl_cp 
            Caption         =   "Code postal"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lbl_adr2 
            Caption         =   "Adresse 2"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lbl_Adr1 
            Caption         =   "Adresse 1"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Fra_Individus 
         Caption         =   "Individus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   25
         Top             =   3840
         Width           =   9255
         Begin VB.CommandButton cmd_SupprimerIndividu 
            Height          =   375
            Left            =   1320
            Picture         =   "frm_DossierGestion.frx":0E0C
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Supprimer l'individu sélectionné"
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmd_CreerIndividu 
            Height          =   375
            Left            =   120
            Picture         =   "frm_DossierGestion.frx":0F0E
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Créer un nouvel individu"
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmd_ModifierIndividu 
            Height          =   375
            Left            =   720
            Picture         =   "frm_DossierGestion.frx":1010
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Modifier l'individu sélectionné"
            Top             =   240
            Width           =   375
         End
         Begin VB.Data Dat_Intermédiaire 
            Caption         =   "Dat_Intermédiaire"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   3240
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   1605
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Data Dat_Dossiers 
            Caption         =   "Dat_Dossiers"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   6120
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "DOSS"
            Top             =   1605
            Visible         =   0   'False
            Width           =   2295
         End
         Begin MSDBGrid.DBGrid grd_Individus 
            Bindings        =   "frm_DossierGestion.frx":1112
            Height          =   1335
            Left            =   120
            OleObjectBlob   =   "frm_DossierGestion.frx":112E
            TabIndex        =   29
            ToolTipText     =   "Liste des individus du foyer demandeur"
            Top             =   720
            Width           =   9015
         End
      End
      Begin VB.Frame fra_Outils 
         Height          =   600
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   9255
         Begin VB.CommandButton cmd_Excel 
            Height          =   375
            Left            =   4560
            Picture         =   "frm_DossierGestion.frx":2023
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Export EXCEL"
            Top             =   150
            Width           =   375
         End
         Begin VB.CommandButton cmd_SupprimerDemande 
            Height          =   375
            Left            =   1320
            Picture         =   "frm_DossierGestion.frx":21AD
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Supprimer la demande sélectionnée"
            Top             =   150
            Width           =   375
         End
         Begin VB.CommandButton cmd_CreerDemande 
            Height          =   375
            Left            =   120
            Picture         =   "frm_DossierGestion.frx":22AF
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Créer une nouvelle demande"
            Top             =   150
            Width           =   375
         End
         Begin VB.CommandButton cmd_ModifierDemande 
            Height          =   375
            Left            =   720
            Picture         =   "frm_DossierGestion.frx":23B1
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Modifier les données de la demande sélectionnée"
            Top             =   150
            Width           =   375
         End
      End
      Begin VB.Frame Fra_Demandes 
         Caption         =   "Liste des demandes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   -74880
         TabIndex        =   11
         Top             =   960
         Width           =   9255
         Begin VB.Data Dat_Demandes 
            Caption         =   "Dat_Demandes"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   315
            Left            =   5640
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   2640
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.CommandButton Cmd_InfosComplementaires 
            Caption         =   "Affic&her les informations complètes ..."
            Height          =   375
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "Permet de visionner les informations relatives à cette demande (souhaits de la famille)"
            Top             =   3240
            Width           =   2775
         End
         Begin VB.CommandButton cmd_Instruction 
            Caption         =   "&Instruction ..."
            Height          =   375
            Left            =   7080
            TabIndex        =   17
            ToolTipText     =   "Accède à la fiche d'instruction de cette demande"
            Top             =   3240
            Width           =   2055
         End
         Begin VB.CommandButton cmdAdoptions 
            Caption         =   "&Suspendre / Reprendre l'instruction ..."
            Height          =   855
            Index           =   0
            Left            =   7080
            TabIndex        =   16
            ToolTipText     =   "Suspendre l'instruction de cette demande"
            Top             =   3720
            Width           =   2055
         End
         Begin VB.CommandButton Cmd_RecalculEcheancier 
            Caption         =   "Date de recalcul de l'&échéancier ..."
            Height          =   375
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "Permet de changer la date servant au calcul de l'échéancier de placement"
            Top             =   4200
            Width           =   2775
         End
         Begin VB.CommandButton Cmd_decicomm 
            Caption         =   "Décision commission ..."
            Height          =   375
            Left            =   3600
            TabIndex        =   14
            Top             =   3240
            Width           =   2775
         End
         Begin VB.CommandButton Cmd_ConfDem 
            Caption         =   "Confirmation de la demande ..."
            Height          =   375
            Left            =   3600
            TabIndex        =   13
            Top             =   4200
            Width           =   2775
         End
         Begin VB.CommandButton Cmd_Recours 
            Caption         =   "Recours ..."
            Height          =   375
            Left            =   3600
            TabIndex        =   12
            Top             =   3720
            Width           =   2775
         End
         Begin MSDBGrid.DBGrid grd_Demandes 
            Bindings        =   "frm_DossierGestion.frx":24B3
            Height          =   2895
            Left            =   120
            OleObjectBlob   =   "frm_DossierGestion.frx":24CE
            TabIndex        =   19
            ToolTipText     =   "Liste des demandes pour le dossier en cours"
            Top             =   240
            Width           =   9045
         End
         Begin VB.Line Line1 
            X1              =   6960
            X2              =   6960
            Y1              =   3120
            Y2              =   4680
         End
         Begin VB.Line Line2 
            X1              =   3000
            X2              =   3000
            Y1              =   3120
            Y2              =   4680
         End
      End
      Begin VB.Data Dat_Individus 
         Caption         =   "Dat_Individus"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5400
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSComctlLib.StatusBar stb_Individu 
         Height          =   255
         Left            =   -74880
         TabIndex        =   65
         Top             =   5760
         Width           =   9255
         _ExtentX        =   16325
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
               TextSave        =   "08/07/2008"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               Alignment       =   1
               Object.Width           =   1764
               MinWidth        =   1764
               TextSave        =   "16:41"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   10107
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frm_DossierGestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum IndexCommandButton 'Index des contrôles CommandButton
    cmdSUSPENDRE_REPRENDRE = 0
End Enum


Private Sub Cmb_NumDossier_Click()
    Cmd_Actualiser_Click
End Sub

Private Sub Cmd_Actualiser_Click()

Dim RsDossier As Recordset
Dim RsIndi As Recordset
Dim NbMois, Mois, An As Integer
Dim NbIndi As Integer

Dim sel_cant As String, sql As String
Dim sel_commune As String
Dim sel_situationfamil As String
Dim sel_Circo As String
    
  If Cmb_NumDossier.Text <> "" Then
    If Cmb_NumDossier.Text > 0 Then
      MousePointer = vbHourglass
      Dat_dossiers.RecordSource = "SELECT * FROM DOSS WHERE DOSS_NUM = " + Iifch(Cmb_NumDossier.Text)
      Dat_dossiers.Refresh
      If Dat_dossiers.Recordset("DOSS_ADR1") <> "" Then
        Txt_adr1 = Dat_dossiers.Recordset("DOSS_ADR1")
      Else
        Txt_adr1 = ""
      End If
      If Dat_dossiers.Recordset("DOSS_ADR2") <> "" Then
        Txt_adr2 = Dat_dossiers.Recordset("DOSS_ADR2")
      Else
        Txt_adr2 = ""
      End If
      If Dat_dossiers.Recordset("CANT_NUM") <> "" Then
        sel_cant = Dat_dossiers.Recordset("CANT_NUM")
        Dat_Intermédiaire.RecordSource = "SELECT CANT_NOM FROM CANT WHERE CANT_NUM =" + _
                                          sel_cant
        Dat_Intermédiaire.Refresh
        Txt_Canton = Dat_Intermédiaire.Recordset("CANT_NOM")
      Else
        Txt_Canton = ""
      End If
      
      If Dat_dossiers.Recordset("COMM_INSEE") <> "" Then
        sel_commune = Dat_dossiers.Recordset("COMM_INSEE")
        Dat_Intermédiaire.RecordSource = "SELECT COMM_NOM FROM COMM WHERE COMM_INSEE =" + _
                                        sel_commune
        Dat_Intermédiaire.Refresh
        Txt_Commune = Dat_Intermédiaire.Recordset("COMM_NOM")
      Else
        Txt_Commune = ""
      End If
      
      If Dat_dossiers.Recordset("CIRC_NUM") <> "" Then
        sel_Circo = Dat_dossiers.Recordset("CIRC_NUM")
        Dat_Intermédiaire.RecordSource = "SELECT CIRC_NOM FROM CIRC WHERE CIRC_NUM =" + sel_Circo
        Dat_Intermédiaire.Refresh
        Txt_Circonscription = Dat_Intermédiaire.Recordset("CIRC_NOM")
      Else
        Txt_Circonscription = ""
      End If
            
      If Dat_dossiers.Recordset("SITU_DMARIA") <> "" Then
        DTP_DateMariage = Dat_dossiers.Recordset("SITU_DMARIA")
      Else
        DTP_DateMariage = Date
        DTP_DateMariage = ""
      End If
      
      If Dat_dossiers.Recordset("DOSS_CP") <> "" Then
        Txt_CP = Dat_dossiers.Recordset("DOSS_CP")
      Else
        Txt_CP = ""
      End If
      If Dat_dossiers.Recordset("SITU_LMARIA") <> "" Then
        Txt_LieuMariage = Dat_dossiers.Recordset("SITU_LMARIA")
      Else
        Txt_LieuMariage = ""
      End If
      If Dat_dossiers.Recordset("SITF_CODE") <> "" Then
        sel_situationfamil = Dat_dossiers.Recordset("SITF_CODE")
        Dat_Intermédiaire.RecordSource = "SELECT SITF_DESIG FROM SITF WHERE SITF_CODE =" + _
                                        sel_situationfamil
        Dat_Intermédiaire.Refresh
        Txt_SituationFamiliale = Dat_Intermédiaire.Recordset("SITF_DESIG")
      Else
        Txt_SituationFamiliale = ""
      End If
      
      If Dat_dossiers.Recordset("SITU_REGMAT") <> "" Then
        Txt_RegimeMatrimonial = Dat_dossiers.Recordset("SITU_REGMAT")
      Else
        Txt_RegimeMatrimonial = ""
      End If
      
      If Dat_dossiers.Recordset("DOSS_NUMTEL") <> "" Then
        Txt_Tel = Dat_dossiers.Recordset("DOSS_NUMTEL")
      Else
        Txt_Tel = ""
      End If
      If Dat_dossiers.Recordset("DOSS_VILLE") <> "" Then
        Txt_Ville = Dat_dossiers.Recordset("DOSS_VILLE")
      Else
        Txt_Ville = ""
      End If
      
      If Dat_dossiers.Recordset("DOSS_REINS") <> "" Then
        Txt_reinstruction = Dat_dossiers.Recordset("DOSS_REINS")
      Else
        Txt_reinstruction = ""
      End If
      
    'calcul age moyen
      NbIndi = 0
      NbMois = 0
      Mois = 0
      'demandeur
      sql = "select indi_dnais from indi where indi_cdfr=""DE"""
      sql = sql + " and doss_num=" + Iifch(Cmb_NumDossier)
      Set RsIndi = Db.OpenRecordset(sql)
      If Not RsIndi.EOF Then
        NbIndi = NbIndi + 1
        NbMois = Abs(DateDiff("m", Date, RsIndi("indi_dnais")))
      End If
      RsIndi.Close
      
      'conjoint
      sql = "select indi_dnais from indi where indi_cdfr=""CO"""
      sql = sql + " and doss_num=" + Iifch(Cmb_NumDossier)
      Set RsIndi = Db.OpenRecordset(sql)
      If Not RsIndi.EOF Then
        NbIndi = NbIndi + 1
        NbMois = NbMois + Abs(DateDiff("m", Date, RsIndi("indi_dnais")))
      End If
      RsIndi.Close
      
      If NbIndi <> 0 Then
        'calcul An
        An = NbMois / (NbIndi * 12)
        An = Round(An, 0)
        If An * 12 * NbIndi > NbMois Then An = An - 1 'cas arrondi supérieur
        'calcul Mois
        Mois = Int((NbMois - (An * 12 * NbIndi)) / 2)
      End If
      
      'modifie l'enregistrement DOSSIER
      sql = "update doss set doss_an=" + CStr(An) + ",doss_mois=" + CStr(Mois)
      sql = sql + " where doss_num =" + Iifch(Cmb_NumDossier)
      Db.Execute sql, dbFailOnError
    'fin calcul age moyen
      
      If Dat_dossiers.Recordset("DOSS_AN") <> "" Then
        Txt_AnneeMoyenne = Dat_dossiers.Recordset("DOSS_AN")
      Else
        Txt_AnneeMoyenne = ""
      End If
      If Dat_dossiers.Recordset("DOSS_MOIS") <> "" Then
        Txt_MoisMoyenne = Dat_dossiers.Recordset("DOSS_MOIS")
      Else
        Txt_MoisMoyenne = ""
      End If
          
      Dat_Individus.RecordSource = "SELECT * FROM INDI WHERE DOSS_NUM = " + Iifch(Cmb_NumDossier.Text)
      Dat_Individus.Refresh
    End If
    Dat_demandes.Refresh
    MousePointer = vbDefault
  End If
End Sub

Private Sub Cmd_ConfDem_Click()

Dim sql As String

  If Cmb_NumDossier.Text <> "" Then
    If Not Dat_demandes.Recordset.EOF Then
      MousePointer = vbHourglass
      Load Frm_ConfirmationDemandeListe
      Frm_ConfirmationDemandeListe.Cmb_NumDossier.Enabled = False
      Frm_ConfirmationDemandeListe.Cmb_NumDemande.Enabled = False
      Frm_ConfirmationDemandeListe.Txt_NomDemandeur.Enabled = False
      Frm_ConfirmationDemandeListe.Txt_PrenomDemandeur.Enabled = False
      Frm_ConfirmationDemandeListe.DTP_DateDemande.Enabled = False
      
      Frm_ConfirmationDemandeListe.Cmb_NumDossier.ListIndex = PresentDansListBox(Frm_ConfirmationDemandeListe.Cmb_NumDossier, Cmb_NumDossier.Text)
      Frm_ConfirmationDemandeListe.Cmb_NumDemande.ListIndex = PresentDansListBox(Frm_ConfirmationDemandeListe.Cmb_NumDemande, Dat_demandes.Recordset("DEMA_NUM"))
     
      Frm_ConfirmationDemandeListe.DTP_DateDemande = Dat_demandes.Recordset("DEMA_DATE")
      
      Frm_ConfirmationDemandeListe.Show vbModal
      MousePointer = vbDefault
    End If
  End If

End Sub

Private Sub cmd_CreerDemande_Click()
Const pkstrNomProcédure As String = "cmd_CreerDemande_Click"
On Error GoTo Err_cmd_CreerDemande_Click
Dim sql As String
Dim i As Integer

    If Cmb_NumDossier.Text <> "" Then
        stb_GestionDossier.Panels(4) = "CREER DEMANDE"
        Me.MousePointer = vbHourglass
        Load frm_DemandeAdoptionFiche
        frm_DemandeAdoptionFiche.Initialiser comCRÉATION, Cmb_NumDossier.Text
        Me.MousePointer = vbDefault
        frm_DemandeAdoptionFiche.Show vbModal
        stb_GestionDossier.Panels(4).Text = ""
    End If
  stb_GestionDossier.Panels(4) = ""

    Exit Sub
    
Err_cmd_CreerDemande_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub Cmd_CreerDossier_Click()

    stb_GestionDossier.Panels(4) = "CREATION"
    Me.MousePointer = vbHourglass
    Load frm_DossierFiche
    frm_DossierFiche.Txt_num.Enabled = False
    frm_DossierFiche.Txt_num = NouveauCodeNouveauDossier
    frm_DossierFiche.Caption = "Création d'un dossier"
    Me.MousePointer = vbDefault
    frm_DossierFiche.Show vbModal
    Cmd_Actualiser_Click
  
    stb_GestionDossier.Panels(4).Text = ""
    Cmb_NumDossier.Enabled = False
    Dat_Individus.Refresh
    cmd_CreerIndividu.SetFocus
End Sub

Private Sub cmd_CreerIndividu_Click()

    If Cmb_NumDossier.Text <> "" Then
        stb_GestionDossier.Panels(4) = "CREATION INDIV"
        Me.MousePointer = vbHourglass
        Load frm_Individu
        frm_Individu.Caption = "Ajout d'un individu"
        frm_Individu.Txt_NumDossier.Enabled = False
        frm_Individu.Txt_NumDossier = Cmb_NumDossier.Text
        frm_Individu.Txt_NumIndividu.Enabled = False
        frm_Individu.Txt_NumIndividu = NouveauCode("INDI", "INDI_NUM")
        frm_Individu.Txt_Adr = ""
        frm_Individu.Txt_Adr1Pre = ""
        frm_Individu.Txt_Adr2Pre = ""
        frm_Individu.Txt_CPEmployeur = ""
        frm_Individu.Txt_CPpre = ""
        frm_Individu.Txt_DepartementNaissance = ""
        frm_Individu.Txt_EMail = ""
        frm_Individu.Txt_Nom = ""
        frm_Individu.Txt_NomEmployeur = ""
        frm_Individu.Txt_Prenom = ""
        frm_Individu.Txt_TelEmployeur = ""
        frm_Individu.Txt_VilleEmployeur = ""
        frm_Individu.Txt_VilleNaissance = ""
        frm_Individu.Txt_VillePre = ""
        frm_Individu.Cmb_CodeFratrie.ListIndex = 0
        frm_Individu.Cmb_Titre.ListIndex = 0
        frm_Individu.DTP_DateAccueil = Date
        frm_Individu.DTP_DateAccueil = ""
        frm_Individu.DTP_DateDeces = Date
        frm_Individu.DTP_DateDeces = ""
        frm_Individu.DTP_DateDivorce = Date
        frm_Individu.DTP_DateDivorce = ""
        frm_Individu.DTP_DateVeuvage = Date
        frm_Individu.DTP_DateVeuvage = ""
        frm_Individu.DTP_DateNaissance = Date

        Me.MousePointer = vbDefault
        frm_Individu.Show vbModal
        Cmd_Actualiser_Click
        stb_GestionDossier.Panels(4) = ""
    End If
End Sub

Private Sub Cmd_decicomm_Click()
Const pkstrNomProcédure As String = "Cmd_decicomm_Click"
Dim plngNuméroDemande As Long
On Error GoTo Err_Cmd_decicomm_Click

'   Vérifier l'existence de la demande
    plngNuméroDemande = lngNuméroDemandeSélectionnée()
    If plngNuméroDemande = comINDÉFINI Then Exit Sub
    
'   Ouvrir la fenêtre permettant le traitement des commissions
    With frm_DecisionCommission
        .Initialiser plngNuméroDemande
        .Show vbModal
    End With
    
    Dat_demandes.Refresh
    
    Exit Sub
    
Err_Cmd_decicomm_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_Excel_Click()
    
  ExportExcel (Dat_demandes.RecordSource)
  
End Sub

Private Sub Cmd_InfosComplementaires_Click()
Const pkstrNomProcédure As String = "Cmd_InfosComplementaires_Click"
Dim plngNuméroDemande As Long
On Error GoTo Err_Cmd_InfosComplementaires_Click

'   Vérifier l'existence de la demande
    plngNuméroDemande = lngNuméroDemandeSélectionnée()
    If plngNuméroDemande = comINDÉFINI Then Exit Sub
    
'   Ouvrir la fenêtre permettant le traitement de la demande
    With frm_DemandeAdoptionFiche
        stb_GestionDossier.Panels(4) = "INFOS DEMANDE"
        .Initialiser comCONSULTATION, Cmb_NumDossier, plngNuméroDemande
        .Show vbModal
        stb_GestionDossier.Panels(4).Text = ""
    End With
    
    Exit Sub
    
Err_Cmd_InfosComplementaires_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_Instruction_Click()
Const pkstrNomProcédure As String = "cmd_Instruction_Click"
Dim plngNuméroDemande As Long
On Error GoTo Err_cmd_Instruction_Click

'   Vérifier l'existence de la demande
    plngNuméroDemande = lngNuméroDemandeSélectionnée()
    If plngNuméroDemande = comINDÉFINI Then Exit Sub
    
'   Ouvrir la fenêtre permettant le traitement de l'instruction
    With frm_DemandeAdoptionInstruction
        stb_GestionDossier.Panels(4) = "INSTRUCTION"
        .Initialiser plngNuméroDemande
        .Show vbModal
        stb_GestionDossier.Panels(4).Text = ""
    End With
    
    Exit Sub
    
Err_cmd_Instruction_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_Listedossiers_Click()

  If Dat_dossiers.Recordset.RecordCount = 0 Then
    MsgBox "Il n'existe aucun dossier."
  Else
    frm_Dossier.Show vbModal
  End If
    
End Sub

Private Sub cmd_ModifierDemande_Click()
Const pkstrNomProcédure As String = "cmd_ModifierDemande_Click"
Dim plngNuméroDemande As Long
On Error GoTo Err_cmd_ModifierDemande_Click

'   Vérifier l'existence de la demande
    plngNuméroDemande = lngNuméroDemandeSélectionnée()
    If plngNuméroDemande = comINDÉFINI Then Exit Sub
    
'   Ouvrir la fenêtre permettant le traitement de la demande
    With frm_DemandeAdoptionFiche
        stb_GestionDossier.Panels(4) = "MODIF DEMANDE"
        .Initialiser comMODIFICATION, Cmb_NumDossier, plngNuméroDemande
        .Show vbModal
        stb_GestionDossier.Panels(4).Text = ""
    End With
    
    Exit Sub
    
Err_cmd_ModifierDemande_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_ModifierDossier_Click()

    If ((Cmb_NumDossier.ListIndex <> -1) And (Cmb_NumDossier.Text <> "")) Then
        stb_GestionDossier.Panels(4) = "MODIFICATION"
        Me.MousePointer = vbHourglass
        Load frm_DossierFiche
        frm_DossierFiche.Caption = "Modification d'un dossier"
        frm_DossierFiche.Txt_num.Enabled = False
        frm_DossierFiche.Txt_num = Cmb_NumDossier.Text
        frm_DossierFiche.Txt_adr1 = Txt_adr1
        frm_DossierFiche.Txt_adr2 = Txt_adr2
        frm_DossierFiche.Txt_CP = Txt_CP
        frm_DossierFiche.Txt_Departement = "22"
        frm_DossierFiche.Txt_LieuMariage = Txt_LieuMariage
        frm_DossierFiche.Txt_Tel = Txt_Tel
        frm_DossierFiche.Txt_Ville = Txt_Ville
        If DTP_DateMariage <> "" Then
            frm_DossierFiche.DTP_DateMariage = DTP_DateMariage
        Else
            frm_DossierFiche.DTP_DateMariage = Date
            frm_DossierFiche.DTP_DateMariage = ""
        End If
        frm_DossierFiche.Txt_regmat = Txt_RegimeMatrimonial
        frm_DossierFiche.Cmb_Commune.ListIndex = PresentDansListBox(frm_DossierFiche.Cmb_Commune, "<" + CStr(Dat_dossiers.Recordset.Fields("COMM_INSEE")) + ">")
        frm_DossierFiche.Cmb_Canton.ListIndex = PresentDansListBox(frm_DossierFiche.Cmb_Canton, "<" + CStr(Dat_dossiers.Recordset.Fields("CANT_NUM")) + ">")
        frm_DossierFiche.Cmb_Circo.ListIndex = PresentDansListBox(frm_DossierFiche.Cmb_Circo, "<" + CStr(Dat_dossiers.Recordset.Fields("CIRC_NUM")) + ">")
        
        frm_DossierFiche.Cmb_SituationFamiliale.ListIndex = PresentDansListBox(frm_DossierFiche.Cmb_SituationFamiliale, "<" + CStr(Dat_dossiers.Recordset.Fields("SITF_CODE")) + ">")
        frm_DossierFiche.Cmb_reins.ListIndex = PresentDansListBox(frm_DossierFiche.Cmb_reins, Txt_reinstruction)
        
        Me.MousePointer = vbDefault
        frm_DossierFiche.Show vbModal
        Cmd_Actualiser_Click
    End If
  
    stb_GestionDossier.Panels(4).Text = ""
    Dat_Individus.Refresh
End Sub

Private Sub cmd_ModifierIndividu_Click()

  If Cmb_NumDossier.Text <> "" Then
    If Not Dat_Individus.Recordset.EOF Then
      stb_GestionDossier.Panels(4) = "MODIF. INDIVIDU"
      Me.MousePointer = vbHourglass
      Load frm_Individu
      frm_Individu.Caption = "Modification d'un individu"
      frm_Individu.Txt_NumIndividu.Enabled = False
      frm_Individu.Txt_NumDossier.Enabled = False
      frm_Individu.Txt_NumIndividu = Dat_Individus.Recordset("INDI_NUM")
      frm_Individu.Txt_NumDossier = Cmb_NumDossier.Text
      If Dat_Individus.Recordset("TRAV_ADDEMP") <> "" Then
          frm_Individu.Txt_Adr = Dat_Individus.Recordset("TRAV_ADDEMP")
      Else
          frm_Individu.Txt_Adr = ""
      End If
      If Dat_Individus.Recordset("INDI_ADRPRE1") <> "" Then
          frm_Individu.Txt_Adr1Pre = Dat_Individus.Recordset("INDI_ADRPRE1")
      Else
          frm_Individu.Txt_Adr1Pre = ""
      End If
      If Dat_Individus.Recordset("INDI_ADRPRE2") <> "" Then
          frm_Individu.Txt_Adr2Pre = Dat_Individus.Recordset("INDI_ADRPRE2")
      Else
          frm_Individu.Txt_Adr2Pre = ""
      End If
      If Dat_Individus.Recordset("TRAV_CPEMP") <> "" Then
          frm_Individu.Txt_CPEmployeur = Dat_Individus.Recordset("TRAV_CPEMP")
      Else
          frm_Individu.Txt_CPEmployeur = ""
      End If
      If Dat_Individus.Recordset("INDI_CODPRE") <> "" Then
          frm_Individu.Txt_CPpre = Dat_Individus.Recordset("INDI_CODPRE")
      Else
          frm_Individu.Txt_CPpre = ""
      End If
      If Dat_Individus.Recordset("INDI_DPNAI") <> "" Then
          frm_Individu.Txt_DepartementNaissance = Dat_Individus.Recordset("INDI_DPNAI")
      Else
          frm_Individu.Txt_DepartementNaissance = ""
      End If
      If Dat_Individus.Recordset("INDI_NOM") <> "" Then
          frm_Individu.Txt_Nom = Dat_Individus.Recordset("INDI_NOM")
      Else
          frm_Individu.Txt_Nom = ""
      End If
      If Dat_Individus.Recordset("TRAV_NOMEMP") <> "" Then
          frm_Individu.Txt_NomEmployeur = Dat_Individus.Recordset("TRAV_NOMEMP")
      Else
          frm_Individu.Txt_NomEmployeur = ""
      End If
      If Dat_Individus.Recordset("INDI_PREN") <> "" Then
          frm_Individu.Txt_Prenom = Dat_Individus.Recordset("INDI_PREN")
      Else
          frm_Individu.Txt_Prenom = ""
      End If
      If Dat_Individus.Recordset("TRAV_TELEMP") <> "" Then
          frm_Individu.Txt_TelEmployeur = Dat_Individus.Recordset("TRAV_TELEMP")
      Else
          frm_Individu.Txt_TelEmployeur = ""
      End If
      If Dat_Individus.Recordset("TRAV_VILLEMP") <> "" Then
          frm_Individu.Txt_VilleEmployeur = Dat_Individus.Recordset("TRAV_VILLEMP")
      Else
          frm_Individu.Txt_VilleEmployeur = ""
      End If
      If Dat_Individus.Recordset("INDI_LNAIS") <> "" Then
          frm_Individu.Txt_VilleNaissance = Dat_Individus.Recordset("INDI_LNAIS")
      Else
          frm_Individu.Txt_VilleNaissance = ""
      End If
      If Dat_Individus.Recordset("INDI_VILPRE") <> "" Then
          frm_Individu.Txt_VillePre = Dat_Individus.Recordset("INDI_VILPRE")
      Else
          frm_Individu.Txt_VillePre = ""
      End If
      If Dat_Individus.Recordset("ASMA_CODE") <> "" Then
          frm_Individu.Cmb_AssuranceMaladie.ListIndex = PresentDansListBox(frm_Individu.Cmb_AssuranceMaladie, "<" + CStr(Dat_Individus.Recordset("ASMA_CODE")) + ">")
      End If
      If Dat_Individus.Recordset("INDI_CDFR") <> "" Then
          frm_Individu.Cmb_CodeFratrie.ListIndex = PresentDansListBox(frm_Individu.Cmb_CodeFratrie, "<" + CStr(Dat_Individus.Recordset("INDI_CDFR")) + ">")
      Else
          frm_Individu.Cmb_CodeFratrie.Text = ""
      End If
      If Dat_Individus.Recordset("INDI_NATIO") <> "" Then
          frm_Individu.txt_nationalite.Text = Dat_Individus.Recordset("INDI_NATIO")
      End If
      If Dat_Individus.Recordset("PROF_CODE") <> "" Then
          frm_Individu.Cmb_Profession.ListIndex = PresentDansListBox(frm_Individu.Cmb_Profession, "<" + CStr(Dat_Individus.Recordset("PROF_CODE")) + ">")
      End If
      If Dat_Individus.Recordset("INDI_TITR") <> "" Then
          frm_Individu.Cmb_Titre.Text = Dat_Individus.Recordset("INDI_TITR")
      Else
          frm_Individu.Cmb_Titre.Text = ""
      End If
      If Dat_Individus.Recordset("INDI_EMAIL") <> "" Then
          frm_Individu.Txt_EMail = Dat_Individus.Recordset("INDI_EMAIL")
      Else
          frm_Individu.Txt_EMail = ""
      End If
      If Dat_Individus.Recordset("INDI_DACCU") <> "" Then
          frm_Individu.DTP_DateAccueil = Dat_Individus.Recordset("INDI_DACCU")
      Else
          frm_Individu.DTP_DateAccueil = Date
          frm_Individu.DTP_DateAccueil = ""
      End If
      If Dat_Individus.Recordset("INDI_DDECE") <> "" Then
          frm_Individu.DTP_DateDeces = Dat_Individus.Recordset("INDI_DDECE")
      Else
          frm_Individu.DTP_DateDeces = Date
          frm_Individu.DTP_DateDeces = ""
      End If
      If Dat_Individus.Recordset("INDI_DDIVO") <> "" Then
          frm_Individu.DTP_DateDivorce = Dat_Individus.Recordset("INDI_DDIVO")
      Else
          frm_Individu.DTP_DateDivorce = Date
          frm_Individu.DTP_DateDivorce = ""
      End If
      If Dat_Individus.Recordset("INDI_DVEUV") <> "" Then
          frm_Individu.DTP_DateVeuvage = Dat_Individus.Recordset("INDI_DVEUV")
      Else
          frm_Individu.DTP_DateVeuvage = Date
          frm_Individu.DTP_DateVeuvage = ""
      End If
      If Dat_Individus.Recordset("INDI_DNAIS") <> "" Then
          frm_Individu.DTP_DateNaissance = Dat_Individus.Recordset("INDI_DNAIS")
      Else
          frm_Individu.DTP_DateNaissance = Date
          frm_Individu.DTP_DateNaissance = ""
      End If
      
      Me.MousePointer = vbDefault
      frm_Individu.Show vbModal
      Cmd_Actualiser_Click
      stb_GestionDossier.Panels(4) = ""
    End If
  End If
  
End Sub

Private Sub cmd_Quitter_Click()

    Dat_Intermédiaire.RecordSource = "SELECT COUNT(INDI_NUM) FROM INDI WHERE INDI_CDFR = ""DE"" AND DOSS_NUM = " + Iifch(Cmb_NumDossier.Text)
    Dat_Intermédiaire.Refresh
    Dat_dossiers.Refresh
    If (((Dat_Intermédiaire.Recordset.Fields(0) <> 0) And (Not Dat_dossiers.Recordset.EOF)) Or (Cmb_NumDossier.Text = "")) Then
        Unload Me
    Else
        MsgBox "Il n'y a aucun individu DEMANDEUR pour ce dossier, vous devez en saisir un.", vbOKOnly, "ATTENTION !"
    End If
End Sub

Private Sub Cmd_RecalculEcheancier_Click()
Const pkstrNomProcédure As String = "Cmd_RecalculEcheancier_Click"
Dim plngNuméroDemande As Long
On Error GoTo Err_Cmd_RecalculEcheancier_Click

'   Vérifier l'existence de la demande
    plngNuméroDemande = lngNuméroDemandeSélectionnée()
    If plngNuméroDemande = comINDÉFINI Then Exit Sub
    
'   Ouvrir la fenêtre permettant le recalcul de l'échéancier
    With frm_RecalculEcheancier
        .Initialiser plngNuméroDemande
        .Show vbModal
    End With
    
    Exit Sub
    
Err_Cmd_RecalculEcheancier_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub Cmd_Recours_Click()
Const pkstrNomProcédure As String = "Cmd_Recours_Click"
Dim plngNuméroDemande As Long
On Error GoTo Err_Cmd_Recours_Click

'   Vérifier l'existence de la demande
    plngNuméroDemande = lngNuméroDemandeSélectionnée()
    If plngNuméroDemande = comINDÉFINI Then Exit Sub
    
'   Ouvrir la fenêtre permettant le traitement des recours
    With Frm_Recours
        .Initialiser plngNuméroDemande
        .Show vbModal
    End With
    
    Exit Sub
    
Err_Cmd_Recours_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_SupprimerDemande_Click()
Const pkstrNomProcédure As String = "cmd_SupprimerDemande_Click"
Dim Rs As Recordset
Dim sql As String
On Error GoTo Err_cmd_SupprimerDemande_Click

  stb_GestionDossier.Panels(4) = "SUPPR DEMANDE"
  If Cmb_NumDossier.Text <> "" Then
    If Not Dat_demandes.Recordset.EOF Then
      sql = "select count(plac_num) from plac where dema_num=" & Dat_demandes.Recordset("dema_num")
      Set Rs = Db.OpenRecordset(sql)
      If Rs.Fields(0) <> 0 Then
        MsgBox "Vous ne pouvez pas supprimer cette demande." + Chr(13) + "Un placement a été réalisé.", vbExclamation, "ERREUR"
      Else
        MousePointer = vbHourglass
        Load frm_DemandeAdoptionFiche
        frm_DemandeAdoptionFiche.Initialiser comSUPPRESSION, Cmb_NumDossier.Text, Dat_demandes.Recordset("DEMA_NUM")
        MousePointer = vbDefault
        frm_DemandeAdoptionFiche.Show vbModal
        stb_GestionDossier.Panels(4).Text = ""
      End If
      Rs.Close
    End If
  End If
  stb_GestionDossier.Panels(4) = ""

    Exit Sub
    
Err_cmd_SupprimerDemande_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_SupprimerDossier_Click()

Dim Rs As Recordset
Dim sql As String

  If Cmb_NumDossier.Text <> "" Then
    sql = "select count(plac_num) from plac where dema_num in"
    sql = sql + "(select dema_num from dema where doss_num=" & Iifch(Cmb_NumDossier.Text) & ")"
    Set Rs = Db.OpenRecordset(sql)
    If Rs.Fields(0) <> 0 Then
      MsgBox "Vous ne pouvez pas supprimer ce dossier." + Chr(13) + "Il existe un placement dans ce dossier.", vbExclamation, "ERREUR"
    Else
      stb_GestionDossier.Panels(4) = "SUPPRESSION"
      MousePointer = vbHourglass
      Load frm_DossierFiche
      frm_DossierFiche.Caption = "Suppression d'un dossier"
      frm_DossierFiche.Txt_num.Enabled = False
      frm_DossierFiche.Txt_adr1.Enabled = False
      frm_DossierFiche.Txt_adr2.Enabled = False
      frm_DossierFiche.Txt_CP.Enabled = False
      frm_DossierFiche.Txt_LieuMariage.Enabled = False
      frm_DossierFiche.Txt_Tel.Enabled = False
      frm_DossierFiche.Txt_Ville.Enabled = False
      frm_DossierFiche.Cmb_Canton.Enabled = False
      frm_DossierFiche.Cmb_Circo.Enabled = False
      frm_DossierFiche.Cmb_Commune.Enabled = False
      frm_DossierFiche.Txt_regmat.Enabled = False
      frm_DossierFiche.Cmb_reins.Enabled = False
      frm_DossierFiche.Cmb_SituationFamiliale.Enabled = False
      frm_DossierFiche.DTP_DateMariage.Enabled = False

      frm_DossierFiche.Txt_num = Cmb_NumDossier.Text
      frm_DossierFiche.Txt_adr1 = Txt_adr1
      frm_DossierFiche.Txt_adr2 = Txt_adr2
      frm_DossierFiche.Txt_CP = Txt_CP
      frm_DossierFiche.Txt_Departement = "22"
      frm_DossierFiche.Txt_LieuMariage = Txt_LieuMariage
      frm_DossierFiche.Txt_Tel = Txt_Tel
      frm_DossierFiche.Txt_Ville = Txt_Ville
      frm_DossierFiche.Cmb_Commune.ListIndex = PresentDansListBox(frm_DossierFiche.Cmb_Commune, "<" + CStr(frm_DossierGestion.Dat_dossiers.Recordset.Fields("COMM_INSEE")) + ">")
      frm_DossierFiche.Cmb_Canton.ListIndex = PresentDansListBox(frm_DossierFiche.Cmb_Canton, "<" + CStr(frm_DossierGestion.Dat_dossiers.Recordset.Fields("CANT_NUM")) + ">")
      frm_DossierFiche.Txt_regmat.Text = Txt_RegimeMatrimonial
      frm_DossierFiche.Cmb_reins.Text = Txt_reinstruction
      frm_DossierFiche.Cmb_SituationFamiliale.ListIndex = PresentDansListBox(frm_DossierFiche.Cmb_SituationFamiliale, "<" + CStr(frm_DossierGestion.Dat_dossiers.Recordset.Fields("SITF_CODE")) + ">")
      If DTP_DateMariage <> "" Then
          frm_DossierFiche.DTP_DateMariage = DTP_DateMariage
      Else
          frm_DossierFiche.DTP_DateMariage = Date
          frm_DossierFiche.DTP_DateMariage = ""
      End If
      
      MousePointer = vbDefault
      frm_DossierFiche.Show vbModal
    End If
    Rs.Close
    stb_GestionDossier.Panels(4).Text = ""
    Dat_Individus.Refresh
    
  End If
  
End Sub

Private Sub cmd_SupprimerIndividu_Click()

    If Cmb_NumDossier.Text <> "" Then
        If Not Dat_Individus.Recordset.EOF Then
            stb_GestionDossier.Panels(4) = "SUPPR. INDIVIDU"
            Me.MousePointer = vbHourglass
            Load frm_Individu
            frm_Individu.Caption = "Suppression d'un individu"
            
            frm_Individu.Txt_Adr.Enabled = False
            frm_Individu.Txt_Adr1Pre.Enabled = False
            frm_Individu.Txt_Adr2Pre.Enabled = False
            frm_Individu.Txt_CPEmployeur.Enabled = False
            frm_Individu.Txt_CPpre.Enabled = False
            frm_Individu.Txt_DepartementNaissance.Enabled = False
            frm_Individu.Txt_Nom.Enabled = False
            frm_Individu.Txt_NomEmployeur.Enabled = False
            frm_Individu.Txt_NumDossier.Enabled = False
            frm_Individu.Txt_NumIndividu.Enabled = False
            frm_Individu.Txt_Prenom.Enabled = False
            frm_Individu.Txt_TelEmployeur.Enabled = False
            frm_Individu.Txt_VilleEmployeur.Enabled = False
            frm_Individu.Txt_VilleNaissance.Enabled = False
            frm_Individu.Txt_VillePre.Enabled = False
            frm_Individu.Txt_EMail.Enabled = False
            frm_Individu.Cmb_AssuranceMaladie.Enabled = False
            frm_Individu.Cmb_CodeFratrie.Enabled = False
            frm_Individu.txt_nationalite.Enabled = False
            frm_Individu.Cmb_Profession.Enabled = False
            frm_Individu.Cmb_Titre.Enabled = False
            frm_Individu.DTP_DateNaissance.Enabled = False
            frm_Individu.DTP_DateAccueil.Enabled = False
            frm_Individu.DTP_DateDeces.Enabled = False
            frm_Individu.DTP_DateDivorce.Enabled = False
            frm_Individu.DTP_DateVeuvage.Enabled = False
            
            frm_Individu.Txt_NumDossier = Cmb_NumDossier.Text
            frm_Individu.Txt_NumIndividu = Dat_Individus.Recordset("INDI_NUM")
            If Dat_Individus.Recordset("TRAV_ADDEMP") <> "" Then
                frm_Individu.Txt_Adr = Dat_Individus.Recordset("TRAV_ADDEMP")
            Else
                frm_Individu.Txt_Adr = ""
            End If
            If Dat_Individus.Recordset("INDI_ADRPRE1") <> "" Then
                frm_Individu.Txt_Adr1Pre = Dat_Individus.Recordset("INDI_ADRPRE1")
            Else
                frm_Individu.Txt_Adr1Pre = ""
            End If
            If Dat_Individus.Recordset("INDI_ADRPRE2") <> "" Then
                frm_Individu.Txt_Adr2Pre = Dat_Individus.Recordset("INDI_ADRPRE2")
            Else
                frm_Individu.Txt_Adr2Pre = ""
            End If
            If Dat_Individus.Recordset("TRAV_CPEMP") <> "" Then
                frm_Individu.Txt_CPEmployeur = Dat_Individus.Recordset("TRAV_CPEMP")
            Else
                frm_Individu.Txt_CPEmployeur = ""
            End If
            If Dat_Individus.Recordset("INDI_CODPRE") <> "" Then
                frm_Individu.Txt_CPpre = Dat_Individus.Recordset("INDI_CODPRE")
            Else
                frm_Individu.Txt_CPpre = ""
            End If
            If Dat_Individus.Recordset("INDI_DPNAI") <> "" Then
                frm_Individu.Txt_DepartementNaissance = Dat_Individus.Recordset("INDI_DPNAI")
            Else
                frm_Individu.Txt_DepartementNaissance = ""
            End If
            If Dat_Individus.Recordset("INDI_NOM") <> "" Then
                frm_Individu.Txt_Nom = Dat_Individus.Recordset("INDI_NOM")
            Else
                frm_Individu.Txt_Nom = ""
            End If
            If Dat_Individus.Recordset("TRAV_NOMEMP") <> "" Then
                frm_Individu.Txt_NomEmployeur = Dat_Individus.Recordset("TRAV_NOMEMP")
            Else
                frm_Individu.Txt_NomEmployeur = ""
            End If
            If Dat_Individus.Recordset("INDI_PREN") <> "" Then
                frm_Individu.Txt_Prenom = Dat_Individus.Recordset("INDI_PREN")
            Else
                frm_Individu.Txt_Prenom = ""
            End If
            If Dat_Individus.Recordset("INDI_EMAIL") <> "" Then
                frm_Individu.Txt_EMail = Dat_Individus.Recordset("INDI_EMAIL")
            Else
                frm_Individu.Txt_EMail = ""
            End If
            If Dat_Individus.Recordset("TRAV_TELEMP") <> "" Then
                frm_Individu.Txt_TelEmployeur = Dat_Individus.Recordset("TRAV_TELEMP")
            Else
                frm_Individu.Txt_TelEmployeur = ""
            End If
            If Dat_Individus.Recordset("TRAV_VILLEMP") <> "" Then
                frm_Individu.Txt_VilleEmployeur = Dat_Individus.Recordset("TRAV_VILLEMP")
            Else
                frm_Individu.Txt_VilleEmployeur = ""
            End If
            If Dat_Individus.Recordset("INDI_LNAIS") <> "" Then
                frm_Individu.Txt_VilleNaissance = Dat_Individus.Recordset("INDI_LNAIS")
            Else
                frm_Individu.Txt_VilleNaissance = ""
            End If
            If Dat_Individus.Recordset("INDI_VILPRE") <> "" Then
                frm_Individu.Txt_VillePre = Dat_Individus.Recordset("INDI_VILPRE")
            Else
                frm_Individu.Txt_VillePre = ""
            End If
            If Dat_Individus.Recordset("ASMA_CODE") <> "" Then
                frm_Individu.Cmb_AssuranceMaladie.ListIndex = PresentDansListBox(frm_Individu.Cmb_AssuranceMaladie, "<" + CStr(Dat_Individus.Recordset("ASMA_CODE")) + ">")
            End If
            If Dat_Individus.Recordset("INDI_CDFR") <> "" Then
                frm_Individu.Cmb_CodeFratrie.ListIndex = PresentDansListBox(frm_Individu.Cmb_CodeFratrie, "<" + CStr(Dat_Individus.Recordset("INDI_CDFR")) + ">")
            Else
                frm_Individu.Cmb_CodeFratrie.Text = ""
            End If
            If Dat_Individus.Recordset("INDI_NATIO") <> "" Then
                frm_Individu.txt_nationalite.Text = Dat_Individus.Recordset("INDI_NATIO")
            End If
            If Dat_Individus.Recordset("PROF_CODE") <> "" Then
                frm_Individu.Cmb_Profession.ListIndex = PresentDansListBox(frm_Individu.Cmb_Profession, "<" + CStr(Dat_Individus.Recordset("PROF_CODE")) + ">")
            End If
            If Dat_Individus.Recordset("INDI_TITR") <> "" Then
                frm_Individu.Cmb_Titre.Text = Dat_Individus.Recordset("INDI_TITR")
            Else
                frm_Individu.Cmb_Titre.Text = ""
            End If
            If Dat_Individus.Recordset("INDI_DACCU") <> "" Then
                frm_Individu.DTP_DateAccueil = Dat_Individus.Recordset("INDI_DACCU")
            Else
                frm_Individu.DTP_DateAccueil = ""
            End If
            If Dat_Individus.Recordset("INDI_DDECE") <> "" Then
                frm_Individu.DTP_DateDeces = Dat_Individus.Recordset("INDI_DDECE")
            Else
                frm_Individu.DTP_DateDeces = ""
            End If
            If Dat_Individus.Recordset("INDI_DDIVO") <> "" Then
                frm_Individu.DTP_DateDivorce = Dat_Individus.Recordset("INDI_DDIVO")
            Else
                frm_Individu.DTP_DateDivorce = ""
            End If
            If Dat_Individus.Recordset("INDI_DVEUV") <> "" Then
                frm_Individu.DTP_DateVeuvage = Dat_Individus.Recordset("INDI_DVEUV")
            Else
                frm_Individu.DTP_DateVeuvage = ""
            End If
            
            Me.MousePointer = vbDefault
            frm_Individu.Show vbModal
            Cmd_Actualiser_Click
            stb_GestionDossier.Panels(4) = ""
        End If
    End If
End Sub

Private Sub cmdAdoptions_Click(Index As Integer)
Const pkstrNomProcédure As String = "cmdAdoptions_Click"
Dim plngNuméroDemande As Long
On Error GoTo Err_cmdAdoptions_Click

    Select Case Index
    
'       Suspendre / Reprendre l'instruction
        Case cmdSUSPENDRE_REPRENDRE
        
'           Vérifier l'existence de la demande
            plngNuméroDemande = lngNuméroDemandeSélectionnée()
            If plngNuméroDemande = comINDÉFINI Then Exit Sub
            
'           Ouvrir la fenêtre permettant le traitement de la suspension ou de la reprise
            With frm_DemandeAdoptionSuspension
                .Initialiser plngNuméroDemande
                .Show vbModal
            End With

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

Private Sub Form_Load()

    Dat_Individus.DatabaseName = CheminBase
    Dat_dossiers.DatabaseName = CheminBase
    Dat_demandes.DatabaseName = CheminBase
    Dat_Intermédiaire.DatabaseName = CheminBase
    
    Txt_adr1.Enabled = False
    Txt_adr2.Enabled = False
    Txt_Canton.Enabled = False
    Txt_Commune.Enabled = False
    Txt_Circonscription.Enabled = False
    Txt_CP.Enabled = False
    Txt_LieuMariage.Enabled = False
    Txt_RegimeMatrimonial.Enabled = False
    Txt_reinstruction.Enabled = False
    Txt_SituationFamiliale.Enabled = False
    Txt_Tel.Enabled = False
    Txt_Ville.Enabled = False
    DTP_DateMariage.Enabled = False


    Txt_Departement = "22"
    Call ListeDossiers(Dat_dossiers, Cmb_NumDossier)
'On verifie que la combo possède au moins un enregistrement afin
'de pouvoir afficher le premier enregistrement dans cette même combo
    If Cmb_NumDossier.ListIndex > 0 Then
        Cmb_NumDossier.ListIndex = 0
    End If
End Sub

Private Sub grd_Demandes_HeadClick(ByVal ColIndex As Integer)
    Dim colonne As String, ClauseOrder As String
    
    colonne = grd_Demandes.Columns(ColIndex).DataField
    If ClauseOrder = " ORDER BY " + colonne + " ASC" Then
        ClauseOrder = " ORDER BY " + colonne + " DESC"
    Else
        ClauseOrder = " ORDER BY " + colonne + " ASC"
    End If
    If InStr(Dat_demandes.RecordSource, " ORDER BY ") <> 0 Then
        Dat_demandes.RecordSource = Left$(Dat_demandes.RecordSource, InStr(Dat_demandes.RecordSource, "ORDER BY") - 1) + ClauseOrder
    Else
        Dat_demandes.RecordSource = Dat_demandes.RecordSource + ClauseOrder
    End If
    Dat_demandes.Refresh
End Sub

Private Sub grd_Individus_HeadClick(ByVal ColIndex As Integer)
    Dim colonne As String, ClauseOrder As String
    
    colonne = grd_Individus.Columns(ColIndex).DataField
    If ClauseOrder = " ORDER BY " + colonne + " ASC" Then
        ClauseOrder = " ORDER BY " + colonne + " DESC"
    Else
        ClauseOrder = " ORDER BY " + colonne + " ASC"
    End If
    If InStr(Dat_Individus.RecordSource, " ORDER BY ") <> 0 Then
        Dat_Individus.RecordSource = Left$(Dat_Individus.RecordSource, InStr(Dat_Individus.RecordSource, "ORDER BY") - 1) + ClauseOrder
    Else
        Dat_Individus.RecordSource = Dat_Individus.RecordSource + ClauseOrder
    End If
    Dat_Individus.Refresh
End Sub

Private Sub SST_Dossier_Click(PreviousTab As Integer)
Const pkstrNomProcédure As String = "SST_Dossier_Click"
On Error GoTo Err_SST_Dossier_Click

    If PreviousTab = 0 Then
    
'       Si le numéro de dossier n'est pas saisi
        If Cmb_NumDossier = "" Then
'           Refuser la navigation sur d'autres onglets que le principal
            SST_Dossier.Tab = 0
            Err.Raise mesSÉLECTION_MANQUANTE
            Exit Sub
        End If
    
        Dat_Intermédiaire.RecordSource = "SELECT COUNT(INDI_NUM) FROM INDI WHERE INDI_CDFR = ""DE"" AND DOSS_NUM = " + Iifch(Cmb_NumDossier.Text)
        Dat_Intermédiaire.Refresh
        Dat_dossiers.Refresh
        If Not (((Dat_Intermédiaire.Recordset.Fields(0) <> 0) And (Not Dat_dossiers.Recordset.EOF)) Or (Cmb_NumDossier.Text = "")) Then
            MsgBox "Il n'y a aucun individu DEMANDEUR pour ce dossier, vous devez en saisir un.", vbOKOnly, "ATTENTION !"
            SST_Dossier.Tab = 0
            Exit Sub
        End If

        Cmd_CreerDossier.Enabled = False
        cmd_ModifierDossier.Enabled = False
        cmd_SupprimerDossier.Enabled = False
        Cmd_Actualiser.Enabled = False
        Cmb_NumDossier.Enabled = False
        cmd_Listedossiers.Enabled = False
        
        If Cmb_NumDossier.Text <> "" Then
            Fra_Demandes.Caption = "Liste des demandes pour le dossier numéro " + Cmb_NumDossier.Text
            Dat_demandes.RecordSource = "SELECT DEMA.*, type_libelle, natu_libelle FROM DEMA, TYPE, NATU" + _
                " WHERE DOSS_NUM = " + Iifch(Cmb_NumDossier.Text) + _
                " AND TYPE.TYPE_CODE = DEMA.TYPE_CODE" + _
                " AND NATU.NATU_CODE = DEMA.NATU_CODE ORDER BY DEMA_NUM"
            Dat_demandes.Refresh
        End If
        
    Else
        Cmd_CreerDossier.Enabled = True
        cmd_ModifierDossier.Enabled = True
        cmd_SupprimerDossier.Enabled = True
        Cmd_Actualiser.Enabled = True
        Cmb_NumDossier.Enabled = True
        cmd_Listedossiers.Enabled = True
    End If
    
    Exit Sub
    
Err_SST_Dossier_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:             pintTraitement = mesINTERROMPRE
        Case mesSÉLECTION_MANQUANTE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Cmb_NumDossier, "d'un numéro de dossier", "Fonctionnalite refusée")
        Case Else:                      pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Retourner le numéro de la demande sélectionnée
'   Note: La valeur comINDÉFINI est retournée si aucune demande n'est sélectionnée

Private Function lngNuméroDemandeSélectionnée() As Long
Const pkstrNomProcédure As String = "lngNuméroDemandeSélectionnée"
Dim pbooDemandeNonCréée As Boolean, pbooDemandeSupprimée As Boolean
On Error GoTo Err_lngNuméroDemandeSélectionnée

    lngNuméroDemandeSélectionnée = comINDÉFINI
    
'   Vérifier l'existence de la demande
    lngNuméroDemandeSélectionnée = Dat_demandes.Recordset("DEMA_NUM")
    If pbooDemandeNonCréée Then Err.Raise mesMESSAGE_013: Exit Function
    If pbooDemandeSupprimée Then Err.Raise mesENREGISTREMENT_SUPPRIMÉ: Exit Function
    
    Exit Function
    
Err_lngNuméroDemandeSélectionnée:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                 pintTraitement = mesINTERROMPRE
        Case daoERREUR_03021:               pbooDemandeNonCréée = True: Resume Next
        Case daoERREUR_03167:               pbooDemandeSupprimée = True: Resume Next
        Case mesENREGISTREMENT_SUPPRIMÉ:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, "Traitement", "La demande", vbCritical)
        Case mesMESSAGE_013:                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
        Case Else:                          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Function


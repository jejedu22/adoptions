VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Recours 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recours"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   10290
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
      Height          =   2295
      Index           =   0
      Left            =   120
      TabIndex        =   47
      Top             =   6840
      Width           =   10095
      Begin VB.TextBox txtAdoptions 
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
         Height          =   285
         Index           =   1
         Left            =   8760
         TabIndex        =   52
         Text            =   "0 recours"
         ToolTipText     =   "Nombre d'enfants demandés"
         Top             =   540
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdoptions 
         Height          =   375
         Index           =   2
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Supprimer le recours contentieux"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAdoptions 
         Height          =   375
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Modifier le recours contentieux"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAdoptions 
         Height          =   375
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Créer un recours contentieux"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin MSComctlLib.ListView lvwAdoptions 
         Height          =   975
         Index           =   0
         Left            =   240
         TabIndex        =   53
         ToolTipText     =   "Historique des recours contentieux"
         Top             =   960
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   1720
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Déposé le"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Auprès de"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Décision"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Rendue le"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Agrément le"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Motif"
            Object.Width           =   5971
         EndProperty
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Historique des recours contentieux"
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
         TabIndex        =   48
         Top             =   0
         Width           =   2970
      End
   End
   Begin VB.Frame fra_Recours 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   1695
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
         Left            =   3720
         MaxLength       =   50
         TabIndex        =   7
         ToolTipText     =   "Demandeur"
         Top             =   360
         Width           =   6120
      End
      Begin VB.ComboBox cboAdoptions 
         Height          =   315
         Index           =   0
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Date du refus ou du retrait"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox Cmb_NumDossiers 
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
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   11
         ToolTipText     =   "Date de la demande"
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
         Format          =   58654723
         CurrentDate     =   2958465
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<Décision>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   8280
         TabIndex        =   14
         Top             =   1125
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Demandeur"
         Height          =   195
         Index           =   2
         Left            =   2775
         TabIndex        =   6
         Top             =   405
         Width           =   825
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Décision du"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   12
         Top             =   1125
         Width           =   840
      End
      Begin VB.Label Lbl_DateDemande 
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
   Begin MSComctlLib.StatusBar stb_Recours 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   54
      Top             =   9240
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "02/10/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "09:58"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10054
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
   Begin VB.Frame Fra_Admin 
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
      Height          =   3855
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   10095
      Begin VB.TextBox Txt_as 
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
         Left            =   1740
         TabIndex        =   36
         ToolTipText     =   "Assistante sociale"
         Top             =   2640
         Width           =   2295
      End
      Begin VB.CommandButton Cmd_TS 
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
         TabIndex        =   43
         ToolTipText     =   "Sélectionner le psychologue"
         Top             =   3120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton Cmd_TS 
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
         TabIndex        =   37
         ToolTipText     =   "Sélectionner l'assistante sociale"
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.ComboBox Cmb_DecifAdm 
         Height          =   315
         Left            =   8280
         Style           =   2  'Dropdown List
         TabIndex        =   28
         ToolTipText     =   "Décision finale"
         Top             =   930
         Width           =   1575
      End
      Begin VB.ComboBox Cmb_DeciAdm 
         Height          =   315
         Left            =   8280
         Style           =   2  'Dropdown List
         TabIndex        =   26
         ToolTipText     =   "Avis de la commission"
         Top             =   570
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTP_ddeciAdm 
         Height          =   285
         Left            =   2280
         TabIndex        =   22
         ToolTipText     =   "Date de signature de la décision"
         Top             =   930
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58654723
         CurrentDate     =   2958465
      End
      Begin VB.TextBox Txt_RefusAdm 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   30
         ToolTipText     =   "Motif du refus"
         Top             =   1695
         Width           =   7575
      End
      Begin MSComCtl2.DTPicker DTP_RecAdm 
         Height          =   285
         Left            =   8280
         TabIndex        =   18
         ToolTipText     =   "Date de dépôt du recours administratif"
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58654723
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_dcomAdm 
         Height          =   285
         Left            =   2280
         TabIndex        =   20
         ToolTipText     =   "Date de réunion de la commission"
         Top             =   570
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58654723
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_ddemasAdm 
         Height          =   285
         Left            =   4695
         TabIndex        =   38
         ToolTipText     =   "Date de demande du rapport de l'assistante sociale"
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58654723
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_ddempsyAdm 
         Height          =   285
         Left            =   4695
         TabIndex        =   44
         ToolTipText     =   "Date de demande du rapport du psychologue"
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58654723
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_recasAdm 
         Height          =   285
         Left            =   6510
         TabIndex        =   39
         ToolTipText     =   "Date de réception du rapport de l'assistante sociale"
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58654723
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_recpsyAdm 
         Height          =   285
         Left            =   6510
         TabIndex        =   45
         ToolTipText     =   "Date de réception du rapport du psychologue"
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58654723
         CurrentDate     =   2958465
      End
      Begin VB.TextBox Txt_psy 
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
         Left            =   1740
         TabIndex        =   42
         ToolTipText     =   "Psychologue"
         Top             =   3120
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTP_dprrecasAdm 
         Height          =   285
         Left            =   8325
         TabIndex        =   40
         ToolTipText     =   "Date prévue pour la réception du rapport de l'assistante sociale"
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58654723
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_dprrecpsyAdm 
         Height          =   285
         Left            =   8325
         TabIndex        =   46
         ToolTipText     =   "Date prévue pour la réception du rapport du psychologue"
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58654723
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_dpredecAdm 
         Height          =   285
         Left            =   2280
         TabIndex        =   24
         ToolTipText     =   "Date prévue pour la signature de la décision"
         Top             =   1290
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58654723
         CurrentDate     =   2958465
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Recours administratif"
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
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   1800
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date prévue de la décision"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1335
         Width           =   1905
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Réception prévue le"
         Height          =   240
         Left            =   8325
         TabIndex        =   34
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FF0000&
         X1              =   8205
         X2              =   8205
         Y1              =   2280
         Y2              =   3480
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Rapport reçu le"
         Height          =   240
         Left            =   6510
         TabIndex        =   33
         Top             =   2265
         Width           =   1575
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Rapport demandé le"
         Height          =   255
         Left            =   4695
         TabIndex        =   32
         Top             =   2265
         Width           =   1575
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FF0000&
         X1              =   6390
         X2              =   6390
         Y1              =   2280
         Y2              =   3480
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FF0000&
         X1              =   4575
         X2              =   4575
         Y1              =   2280
         Y2              =   3480
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Nom"
         Height          =   255
         Left            =   1740
         TabIndex        =   31
         Top             =   2265
         Width           =   2295
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF0000&
         X1              =   1620
         X2              =   1620
         Y1              =   2280
         Y2              =   3480
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Psychologue"
         Height          =   195
         Left            =   585
         TabIndex        =   41
         Top             =   3150
         Width           =   915
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         X1              =   240
         X2              =   9840
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Assistante sociale"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   2670
         Width           =   1260
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         X1              =   240
         X2              =   9840
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         X1              =   240
         X2              =   9840
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Avis de la commission"
         Height          =   195
         Left            =   6600
         TabIndex        =   25
         Top             =   615
         Width           =   1545
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Décision finale"
         Height          =   195
         Left            =   7110
         TabIndex        =   27
         Top             =   975
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date de la commission"
         Height          =   195
         Left            =   555
         TabIndex        =   19
         Top             =   615
         Width           =   1590
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date de la signature"
         Height          =   195
         Left            =   720
         TabIndex        =   21
         Top             =   975
         Width           =   1425
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Motif du refus"
         Height          =   195
         Left            =   1185
         TabIndex        =   29
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   " Déposé le      "
         Height          =   195
         Left            =   7380
         TabIndex        =   17
         Top             =   30
         Width           =   1005
      End
   End
End
Attribute VB_Name = "Frm_Recours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
                          
Private WithEvents mfrmRecoursContentieux As frmRecoursContentieux
Attribute mfrmRecoursContentieux.VB_VarHelpID = -1
Private WithEvents mfrm_TS As frm_TS
Attribute mfrm_TS.VB_VarHelpID = -1

Private Enum IndexComboBox      'Index des contrôles ComboBox
    cboDÉCISION = 0
End Enum
Private Enum IndexCommandButton 'Index des contrôles CommandButton
    cmdPSYCHOLOGUE = 0
    cmdASSISTANTE_SOCIALE
    cmdCRÉER = 0
    cmdMODIFIER
    cmdSUPPRIMER
End Enum
Private Enum IndexFrame         'Index des contrôles Frame
    fraCONTENTIEUX = 0
End Enum
Private Enum IndexColumnHeader  'Index des contrôles ColumnHeader
    hdrNUMÉRO = 0
    hdrDÉPOSÉ
    hdrAUPRÈS
    hdrDÉCISION
    hdrRENDUE
    hdrMOTIF
End Enum
Private Enum IndexLabel         'Index des contrôles Label
    lblDÉCISION = 4
End Enum
Private Enum IndexListView      'Index des contrôles ListView
    lvwCONTENTIEUX = 0
End Enum
Private Enum IndexTextBox       'Index des contrôles TextBox
    txtDEMANDEUR = 0
    txtCONTENTIEUX
End Enum

Private Const mstrNOMBRE_RECOURS As String = " recours"

Private mvarPsychologueNuméro As Variant, mvarAssistanteSocialeNuméro As Variant
Private mclsDécisions As clsDécisions
Private mrstRecoursAdministratif As Recordset, mrstRecoursContentieux As Recordset
Private mdtmSignature As Date
Private mlngIntervenant As IndexCommandButton
Private mbooInitialisationAvecDemandeDéfinie As Boolean

Private Sub cboAdoptions_Click(Index As Integer)
Const pkstrNomProcédure As String = "cboAdoptions_Click"
Dim pclsDécision As clsDécision
Dim pstrDécision As String
Dim plngDécision As AdoptionsDélibérations
On Error GoTo Err_cboAdoptions_Click

    Select Case Index
    
'       Liste des décisions
        Case cboDÉCISION
        
'           Si elle est définie
            If Not cboAdoptions(cboDÉCISION).ListIndex = 0 Then
            
                Set pclsDécision = mclsDécisions(cboAdoptions(cboDÉCISION).ListIndex)
                pstrDécision = gclsDélibérations(CStr(pclsDécision.lngDécision)).strLibellé
                
'               Implémenter la nature de la décision
                With lblAdoptions(lblDÉCISION)
                    .Visible = True
                    .Caption = pstrDécision
                End With
            
'               Désactiver les identifiants
                modCommun.DéfinirVerrouillageContrôle Cmb_NumDossiers, True, True
                modCommun.DéfinirVerrouillageContrôle Cmb_NumDemande, True, True
                modCommun.DéfinirVerrouillageContrôle cboAdoptions(cboDÉCISION), True, True
            
'               Positionner une liste de DTPicker à la date du jour
                modCommun.PositionnerDateJour DTP_RecAdm, DTP_dcomAdm, DTP_ddeciAdm, DTP_dpredecAdm, DTP_ddemasAdm, DTP_recasAdm, DTP_dprrecasAdm, DTP_ddempsyAdm, DTP_recpsyAdm, DTP_dprrecpsyAdm
                
'               Vérifier l'existence d'un recours administratif
                Set mrstRecoursAdministratif = modDAO.rstSQLExtraction(daoRECOURS_ADMINISTRATIF, True, gstrCONSULTATION_LIMITÉE, dbOpenDynaset, CLng(Cmb_NumDemande), CDate(cboAdoptions(cboDÉCISION)))
                If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
                    modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
                    Exit Sub
                End If
                
'               S'il existe
                If Not modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT Then
                
                    With mrstRecoursAdministratif
                    
'                       Vérifier la cohérence entre les informations présentes sur le recours administratif et celles de l'historique des intervenants
                        modAdoptions.VérifierCohérenceAvecHistoriqueIntervenants CLng(Cmb_NumDemande), 1, AdoptionsCadresInterventions.adpRECOURS_ADMINISTRATIF, ![TRSO_NUMPSY], ![READ_DDEMPSY], ![READ_DPRRECPSY], ![READ_RECPSY], ![TRSO_NUMAS], ![READ_DDEMAS], ![READ_DPRRECAS], ![READ_RECAS]
                    
'                       Positionner les informations relatives au recours administratif
                        DTP_RecAdm = IIf(IsNull(![READ_DDREC]), "", ![READ_DDREC])
                        DTP_dcomAdm = IIf(IsNull(![READ_DCOM]), "", ![READ_DCOM])
                        DTP_ddeciAdm = IIf(IsNull(![READ_DDECI]), "", ![READ_DDECI])
                        DTP_dpredecAdm = IIf(IsNull(![READ_DPREDEC]), "", ![READ_DPREDEC])
                        plngDécision = IIf(IsNull(![READ_DECIC]), AdoptionsDélibérations.adpINDÉTERMINÉ, IIf(![READ_DECIC] = gstrREFUS, adpREFUS, adpACCORD))
                        Cmb_DeciAdm.ListIndex = modCommun.varRecherche(Cmb_DeciAdm, False, plngDécision)
                        plngDécision = IIf(IsNull(![READ_DECIF]), AdoptionsDélibérations.adpINDÉTERMINÉ, IIf(![READ_DECIF] = gstrREFUS, adpREFUS, adpACCORD))
                        Cmb_DecifAdm.ListIndex = modCommun.varRecherche(Cmb_DecifAdm, False, plngDécision)
                        Txt_RefusAdm = modCommun.varValeurFormatéeSiNull(![READ_REFUS], vbString)
                        mvarAssistanteSocialeNuméro = ![TRSO_NUMAS]
                        DTP_ddemasAdm = IIf(IsNull(![READ_DDEMAS]), "", ![READ_DDEMAS])
                        DTP_recasAdm = IIf(IsNull(![READ_RECAS]), "", ![READ_RECAS])
                        DTP_dprrecasAdm = IIf(IsNull(![READ_DPRRECAS]), "", ![READ_DPRRECAS])
                        mvarPsychologueNuméro = ![TRSO_NUMPSY]
                        DTP_ddempsyAdm = IIf(IsNull(![READ_DDEMPSY]), "", ![READ_DDEMPSY])
                        DTP_recpsyAdm = IIf(IsNull(![READ_RECPSY]), "", ![READ_RECPSY])
                        DTP_dprrecpsyAdm = IIf(IsNull(![READ_DPRRECPSY]), "", ![READ_DPRRECPSY])

'                       Si elle est définie, afficher l'identité de l'assistante sociale
                        If Not IsNull(mvarAssistanteSocialeNuméro) Then modAdoptions.AfficherIdentitéIntervenant CInt(mvarAssistanteSocialeNuméro), Txt_as, gstrCONSULTATION_LIMITÉE
'                       S'il est défini, afficher l'identité du psychologue
                        If Not IsNull(mvarPsychologueNuméro) Then modAdoptions.AfficherIdentitéIntervenant CInt(mvarPsychologueNuméro), Txt_psy, gstrCONSULTATION_LIMITÉE
                        
                    End With
                    
'               Sinon
                Else
                
'                   Désélectionner les dates se rapportant au recours administratif
                    Implémenter "<dates du recours administratif>"
                    
                End If

'               Activer les contrôles relatifs au recours administratif
                Fra_Admin.Enabled = True
                If Me.Visible Then DTP_RecAdm.SetFocus
                
'               Activer les contrôles relatifs aux recours contentieux
                fraAdoptions(fraCONTENTIEUX).Enabled = True
                
'               Activer/Désactiver les contrôles se rapportant au recours administratif
                DTP_RecAdm_Change
                
'               Implémenter l'historique des recours contentieux
                Implémenter "Historique des recours contentieux"
                If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub
    
'               Si la date de signature n'est pas saisie
                If IsNull(pclsDécision.varSignature) Then
                    Err.Raise mesDATE_SIGNATURE_ABSENTE
                    modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
'               Sinon
                Else
                    mdtmSignature = pclsDécision.varSignature
                End If
    
            End If
            
    End Select
        
    Exit Sub
    
Err_cboAdoptions_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:             pintTraitement = mesINTERROMPRE
        Case mesDATE_SIGNATURE_ABSENTE: pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, LCase(pstrDécision))
        Case Else:                      pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub Cmb_DecifAdm_Click()
Const pkstrNomProcédure As String = "Cmb_DecifAdm_Click"
Dim pbooVerrouillé As Boolean
On Error GoTo Err_Cmb_DecifAdm_Click

'   Activer/Désactiver le motif selon la nature de la décision finale
    pbooVerrouillé = Not (Cmb_DecifAdm.ItemData(Cmb_DecifAdm.ListIndex) = adpREFUS)
    modCommun.DéfinirVerrouillageContrôle Txt_RefusAdm, pbooVerrouillé
    If pbooVerrouillé Then Txt_RefusAdm = ""

    Exit Sub
    
Err_Cmb_DecifAdm_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
  
End Sub


Private Sub Cmb_NumDemande_Click()
Const pkstrNomProcédure As String = "Cmb_NumDemande_Click"
Dim pclsDécision As clsDécision
Dim prstDateDépôtDemande As Recordset, prstDécisionsRefusRetraits As Recordset
Dim pintItération As Integer
On Error GoTo Err_Cmb_NumDemande_Click

'   Implémenter la date de dépôt de la demande et la date de la commission ayant décidé le refus ou le retrait
    Implémenter "Déposée le,Décision du", "Vrai" & gstrSÉPARATEUR_1 & ""
    
'   Lire la date de dépôt de la demande
    Set prstDateDépôtDemande = modDAO.rstSQLExtraction(daoEXTRACTION_019, True, gstrCONSULTATION_LIMITÉE & gstrSÉPARATEUR_1 & gstrDEMANDE_SUPPRIMÉE, CLng(Cmb_NumDemande))
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
        modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
        Exit Sub
    End If
    DTP_DateDemande = prstDateDépôtDemande![DEMA_DATE]
    
'   Lire les refus et les retraits décidés par les commissions
    Set prstDécisionsRefusRetraits = modDAO.rstSQLExtraction(daoEXTRACTION_016, True, gstrCONSULTATION_LIMITÉE, CLng(Cmb_NumDemande))
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
        modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
        Exit Sub
    End If
    
'   Implémenter la liste des décisions
    With prstDécisionsRefusRetraits
        While Not .EOF
            pintItération = pintItération + 1
            modCommun.ImplémenterListe cboAdoptions(cboDÉCISION), CStr(![u02_dtmHistoCommissionReunion]), pvarNuméro:=pintItération
            Set pclsDécision = mclsDécisions.Ajout(0, ![u02_dtmHistoCommissionReunion], AdoptionsProvenances.adpDEMANDE_INITIALE, ![u02_intHistoCommissionDecision], ![u02_dtmHistoCommissionReunion], ![u02_dtmHistoCommissionSignature], pvarClé:=CStr(pintItération))
            .MoveNext
        Wend
    End With
    
'   S'il n'en existe qu'une, la sélectionner
    If cboAdoptions(cboDÉCISION).ListCount = 2 Then cboAdoptions(cboDÉCISION).ListIndex = 1
    
    If Me.Visible And cboAdoptions(cboDÉCISION).Enabled Then cboAdoptions(cboDÉCISION).SetFocus
    
    Exit Sub
    
Err_Cmb_NumDemande_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


Private Sub Cmb_NumDossiers_Click()
Const pkstrNomProcédure As String = "Cmb_NumDossiers_Click"
On Error GoTo Err_Cmb_NumDossiers_Click

'   Afficher l'identité du demandeur
    modAdoptions.AfficherIdentitéDemandeur Cmb_NumDossiers, txtAdoptions(txtDEMANDEUR), gstrCONSULTATION_LIMITÉE, cmd_Valider
        
'   Implémenter la date de dépôt de la demande et la date de la commission ayant décidé le refus ou le retrait
    Implémenter "Déposée le,Décision du", "Faux" & gstrSÉPARATEUR_1 & ""
    
'   Si la demande n'est pas définie à l'initialisation de la fenêtre
    If Not mbooInitialisationAvecDemandeDéfinie Then
'       Lire les numéros des demandes liées au dossier
        modAdoptions.ImplémenterListe Cmb_NumDemande, "Demandes", gstrCONSULTATION_LIMITÉE, cmd_Valider, Cmb_NumDossiers
'       S'il n'en existe qu'une, la sélectionner
        If Cmb_NumDemande.ListCount = 1 Then Cmb_NumDemande.ListIndex = 0
    End If
    
    If Cmb_NumDemande.Enabled Then Cmb_NumDemande.SetFocus
    
    Exit Sub
    
Err_Cmb_NumDossiers_Click:
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

Private Sub Cmd_TS_Click(Index As Integer)
Const pkstrNomProcédure As String = "Cmd_TS_Click"
On Error GoTo Err_Cmd_TS_Click

'   Ouvrir la fenêtre permettant le traitement de l'intervenant
    With mfrm_TS
        mlngIntervenant = Index
        .Initialiser False, True, True
        .Show vbModal
    End With
    
    Exit Sub
    
Err_Cmd_TS_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_Valider_Click()
Const pkstrNomProcédure As String = "cmd_Valider_Click"
Const pkstrDoublon As String = "Création refusée: Le recours administratif vient d'être créé par un autre utilisateur.", pkstrListeMessages As String = gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_1 & pkstrDoublon
Dim pvarNuméro As Variant, pvarAvis As Variant, pvarDécision As Variant, pvarOriginalDemandé As Variant
Dim pctlFocus As Control
Dim pcmdIntervenant As CommandButton
Dim pdtpDemandé As DTPicker, pdtpReçu As DTPicker, pdtpPrévu As DTPicker
Dim pstrMessages(0) As String, pstrChaîne As String
Dim plngParamètreÉchéancier As DAOParamètresÉchéancier, plngDélibération As AdoptionsDélibérations, plngTraitement As CommunTraitements
Dim pintItération As Integer
On Error GoTo Err_cmd_Valider_Click

'   Forcer si nécessaire l'actualisation de la valeur d'un contrôle DTPicker
    modCommun.GérerActualisationDTPicker Me, cmd_Valider
    
'   Refuser la validation si la date de la décision n'est pas sélectionnée
    If cboAdoptions(cboDÉCISION) = "" Then
        Err.Raise mesSÉLECTION_MANQUANTE
        Exit Sub
    End If
    
'   Pour chaque intervenant
    For pintItération = 1 To 2

        Select Case pintItération
'           Assistante sociale
            Case 1: Set pdtpDemandé = DTP_ddemasAdm:    Set pdtpPrévu = DTP_dprrecasAdm:    plngParamètreÉchéancier = daoRECOURS_ADMINISTRATIF__RÉPONSE_TRAVAILLEUR_SOCIAL
'           Psychologue
            Case 2: Set pdtpDemandé = DTP_ddempsyAdm:   Set pdtpPrévu = DTP_dprrecpsyAdm:   plngParamètreÉchéancier = daoRECOURS_ADMINISTRATIF__RÉPONSE_PSYCHOLOGUE
        End Select
        
'       Si un rapport de l'intervenant est demandé
        If Not IsNull(pdtpDemandé) Then
'           Si la date prévue pour sa réception n'est pas définie
            If IsNull(pdtpPrévu) Then
'               La calculer
                pdtpPrévu = DateAdd("m", modDAO.intParamètreÉchéancier(plngParamètreÉchéancier), pdtpDemandé)
            End If
        End If
    
    Next pintItération
        
'   Ouvrir une transaction
    modDAO.GérerTransaction daoOUVERTURE
    
'   Si la date de dépôt du recours administratif est définie
    If Not IsNull(DTP_RecAdm) Then
    
        plngTraitement = IIf(mrstRecoursAdministratif.RecordCount = 0, comCRÉATION, comMODIFICATION)
    
'       Vérifier la cohérence de la décision
        If modAdoptions.booSaisieIncohérenteDécision(CLng(Cmb_NumDemande), gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_2 & CStr(vbInformation), plngTraitement, CDate(cboAdoptions(cboDÉCISION)), AdoptionsProvenances.adpRECOURS_ADMINISTRATIF, DTP_dcomAdm, Cmb_DecifAdm, DTP_ddeciAdm, pdtpRecoursDépôt:=DTP_RecAdm, pdtmRefusRetraitSignature:=mdtmSignature, pstrRefusRetrait:=LCase(lblAdoptions(lblDÉCISION)), pdtpDécisionPrévue:=DTP_dpredecAdm, pcmdÀDésactiver:=cmd_Valider) Then modDAO.GérerTransaction daoANNULATION: Exit Sub

'       Pour chaque intervenant
        For pintItération = 1 To 2
    
            Select Case pintItération
'               Assistante sociale
                Case 1: pvarNuméro = mvarAssistanteSocialeNuméro:   Set pcmdIntervenant = Cmd_TS(cmdASSISTANTE_SOCIALE):    Set pdtpDemandé = DTP_ddemasAdm:    Set pdtpReçu = DTP_recasAdm:    Set pdtpPrévu = DTP_dprrecasAdm:    pstrChaîne = "de l'assistante sociale": If plngTraitement = comCRÉATION Then pvarOriginalDemandé = Null Else pvarOriginalDemandé = mrstRecoursAdministratif![READ_DDEMAS]
'               Psychologue
                Case 2: pvarNuméro = mvarPsychologueNuméro:         Set pcmdIntervenant = Cmd_TS(cmdPSYCHOLOGUE):           Set pdtpDemandé = DTP_ddempsyAdm:   Set pdtpReçu = DTP_recpsyAdm:   Set pdtpPrévu = DTP_dprrecpsyAdm:   pstrChaîne = "du psychologue":          If plngTraitement = comCRÉATION Then pvarOriginalDemandé = Null Else pvarOriginalDemandé = mrstRecoursAdministratif![READ_DDEMPSY]
            End Select
            
'           Refuser la validation si les informations relatives à l'intervenant sont incohérentes
            If modAdoptions.booSaisieIncohérenteIntervenant(pdtpDemandé, DTP_RecAdm, pvarNuméro, pcmdIntervenant, pdtpDemandé, pdtpReçu, pstrChaîne, "de demande du rapport", "de dépôt du recours administratif", pdtpPrévu:=pdtpPrévu) Then
                modDAO.GérerTransaction daoANNULATION: Exit Sub
            End If
            
'           Si l'intervenant est identifié
            If Not IsNull(pdtpDemandé) Then
            
'               S'il l'était auparavant
                If Not IsNull(pvarOriginalDemandé) Then
                
'                   Mettre à jour l'historique
                    modDAO.SQLMettreÀJour daoMAJ_HISTORIQUE_INTERVENANT, True, gstrVALIDATION_ANNULÉE, CLng(Cmb_NumDemande), (pintItération = 2), pvarOriginalDemandé, 1, pdtpDemandé, pdtpPrévu, pdtpReçu, pvarNuméro
                    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub

'               Sinon
                Else
                
'                   Créer l'historique
                    modDAO.SQLAjouter daoADD_HISTORIQUE_INTERVENANT, True, pkstrListeMessages, CLng(Cmb_NumDemande), AdoptionsCadresInterventions.adpRECOURS_ADMINISTRATIF, (pintItération = 2), pdtpDemandé, 1, pdtpPrévu, pdtpReçu, pvarNuméro
                    If modDAO.lngDernièreErreurDAO() = daoDOUBLON Then modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
                    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                    
                End If
            
'           Sinon
            Else
            
'               S'il l'était auparavant
                If Not IsNull(pvarOriginalDemandé) Then
                
'                   Supprimer l'historique
                    modDAO.SQLSupprimer daoDEL_HISTORIQUE_INTERVENANT, True, gstrVALIDATION_ANNULÉE, CLng(Cmb_NumDemande), (pintItération = 2), pvarOriginalDemandé, 1
                    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                    
                End If
            
            End If
    
        Next pintItération
    
        With mrstRecoursAdministratif
            
'           Si le recours ne doit pas être examiné en commission
            If IsNull(DTP_dcomAdm) Then
            
'               S'il existait un recours administratif à l'origine
                If Not .RecordCount = 0 Then
                
'                   S'il devait être examiné en commission
                    If Not IsNull(![READ_DCOM]) Then
'                       Demander confirmation de la suppression de l'inscription du dossier à la commission
                        Select Case modMessages.lngRéponseUtilisateur(mesSUPPRIMER_INSCRIPTION_COMMISSION, ![READ_DCOM])
'                           L'utilisateur confirme la suppression
                            Case vbYes
'                           L'utilisateur renonce à la suppression
                            Case vbNo
                                DTP_dcomAdm.SetFocus
                                modDAO.GérerTransaction daoANNULATION: Exit Sub
                        End Select
                    End If

                End If
                
            End If

            plngDélibération = Cmb_DeciAdm.ItemData(Cmb_DeciAdm.ListIndex): pvarAvis = IIf(plngDélibération = AdoptionsDélibérations.adpINDÉTERMINÉ, Null, IIf(plngDélibération = adpREFUS, gstrREFUS, gstrACCORD))
            plngDélibération = Cmb_DecifAdm.ItemData(Cmb_DecifAdm.ListIndex): pvarDécision = IIf(plngDélibération = AdoptionsDélibérations.adpINDÉTERMINÉ, Null, IIf(plngDélibération = adpREFUS, gstrREFUS, gstrACCORD))

'           Si un recours administratif est enregistré
            If Not .RecordCount = 0 Then

'               Le mettre à jour
'                   Note: Ne pas utiliser de requête action (de type <mconnODBCAdoptions.Execute pstrChaîneSQL, dbExecDirect>) ou d'instruction SQL, au risque de ne pas être informé des modifications éventuelles apportées par un autre utilisateur sur la date de demande de rapport, et donc de provoquer la création d'un nouvel historique lors de la vérification de la cohérence

                .Edit
                
                ![TRSO_NUMAS] = mvarAssistanteSocialeNuméro
                ![TRSO_NUMPSY] = mvarPsychologueNuméro
                ![READ_DDREC] = DTP_RecAdm
                ![READ_DCOM] = DTP_dcomAdm
                ![READ_DPREDEC] = DTP_dpredecAdm
                ![READ_DDECI] = DTP_ddeciAdm
                ![READ_DECIC] = pvarAvis
                ![READ_DECIF] = pvarDécision
                ![READ_DDEMAS] = DTP_ddemasAdm
                ![READ_DPRRECAS] = DTP_dprrecasAdm
                ![READ_RECAS] = DTP_recasAdm
                ![READ_DDEMPSY] = DTP_ddempsyAdm
                ![READ_DPRRECPSY] = DTP_dprrecpsyAdm
                ![READ_RECPSY] = DTP_recpsyAdm
                ![READ_REFUS] = modCommun.varChaîneFormatée(Txt_RefusAdm, vbString, Null)
                
'               Valider l'opération
                modDAO.ValiderOpérationSurRecordset mrstRecoursAdministratif, pstrListeMessages:=gstrMESSAGE_SUR_VERROUILLAGE_EN_MODIFICATION & gstrSÉPARATEUR_1 & gstrMESSAGE_SUR_SUPPRESSION & gstrSÉPARATEUR_1 & gstrMESSAGE_SUR_MODIFICATION_SANS_CHOIX & gstrSÉPARATEUR_1 & pkstrDoublon, pvarRéponseÀInviteÉcrasement:=vbNo
                Select Case modDAO.lngDernièreErreurDAO()
'                   L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                    Case daoENREGISTREMENT_VERROUILLÉ
                        modDAO.GérerTransaction daoANNULATION: Exit Sub
'                   L'opération a été refusée en raison de la modification par un autre utilisateur des données originales
                    Case dbRecordModified
                        modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
                        modDAO.GérerTransaction daoANNULATION: Exit Sub
'                   L'opération a été refusée en raison de la suppression par un autre utilisateur des données originales
                    Case dbRecordDeleted
                        modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
                        modDAO.GérerTransaction daoANNULATION: Exit Sub
'                   L'opération a généré un conflit inattendu
                    Case Is <> daoAUCUNE
                        modDAO.GérerTransaction daoANNULATION
                        Err.Raise mesCONFLIT_ACCÈS
                End Select

'           Sinon
            Else
                
'               Créer le recours administratif
                modDAO.SQLAjouter daoADD_RECOURS_ADMINISTRATIF, True, pkstrListeMessages, mvarAssistanteSocialeNuméro, mvarPsychologueNuméro, DTP_RecAdm, DTP_dcomAdm, DTP_dpredecAdm, DTP_ddeciAdm, pvarAvis, pvarDécision, DTP_ddemasAdm, DTP_dprrecasAdm, DTP_recasAdm, DTP_ddempsyAdm, DTP_dprrecpsyAdm, DTP_recpsyAdm, modCommun.varChaîneFormatée(Txt_RefusAdm, vbString, Null), CLng(Cmb_NumDemande), CDate(cboAdoptions(cboDÉCISION))
                If modDAO.lngDernièreErreurDAO() = daoDOUBLON Then modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
                If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub

            End If
            
        End With
            
'   Sinon
    Else

        With mrstRecoursAdministratif
        
'           Si aucun recours administratif n'est enregistré
            If .RecordCount = 0 Then
            
                Set pctlFocus = DTP_RecAdm: pstrMessages(0) = "de la date de dépôt du recours administratif"
                Err.Raise mesSAISIE_MANQUANTE
                modDAO.GérerTransaction daoANNULATION: Exit Sub
            
'           Sinon
            Else
            
'               Vérifier la cohérence de la suppression
                If modAdoptions.booSaisieIncohérenteDécision(CLng(Cmb_NumDemande), gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_2 & CStr(vbInformation), comSUPPRESSION, CDate(cboAdoptions(cboDÉCISION)), AdoptionsProvenances.adpRECOURS_ADMINISTRATIF, DTP_dcomAdm, Cmb_DecifAdm, DTP_ddeciAdm, pdtpRecoursDépôt:=DTP_RecAdm, pdtmRefusRetraitSignature:=mdtmSignature, pstrRefusRetrait:=LCase(lblAdoptions(lblDÉCISION)), pdtpDécisionPrévue:=DTP_dpredecAdm, pcmdÀDésactiver:=cmd_Valider) Then modDAO.GérerTransaction daoANNULATION: Exit Sub

'               Demander confirmation de la suppression du recours
                Select Case modMessages.lngRéponseUtilisateur(mesSUPPRIMER, "le recours administratif")
'                   L'utilisateur confirme la suppression
                    Case vbYes
'                   L'utilisateur renonce à la suppression
                    Case vbNo
                        DTP_RecAdm.SetFocus
                        modDAO.GérerTransaction daoANNULATION: Exit Sub
                End Select
                
'               Rafraîchir le recours
'                   Note: Nécessaire pour s'assurer des références des intervenants
                modDAO.ActualiserRecordset mrstRecoursAdministratif, True, gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_2 & CStr(vbInformation)
                If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                
'               Si le recours existe toujours
                If Not modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT Then
                
'                   Pour chaque intervenant
                    For pintItération = 1 To 2
                
                        Select Case pintItération
'                           Assistante sociale
                            Case 1: pvarOriginalDemandé = ![READ_DDEMAS]
'                           Psychologue
                            Case 2: pvarOriginalDemandé = ![READ_DDEMPSY]
                        End Select
                        
'                       S'il est identifié sur le recours
                        If Not IsNull(pvarOriginalDemandé) Then
                    
'                           Supprimer l'historique
                            modDAO.SQLSupprimer daoDEL_HISTORIQUE_INTERVENANT, True, gstrVALIDATION_ANNULÉE, ![DEMA_NUM], (pintItération = 2), pvarOriginalDemandé, 1
                            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                            
                        End If
                        
                    Next pintItération
                
'                   Supprimer le recours
                    modDAO.SQLSupprimer daoDEL_HISTORIQUE_COMMISSION_OU_RECOURS, True, gstrVALIDATION_ANNULÉE, "READ", ![DEMA_NUM], ![u02_dtmHistoCommissionReunion]
                    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                        
                End If
                
            End If
            
        End With
            
    End If
    
'   Actualiser les données relatives à l'agrément
    modAdoptions.ActualiserDonnéesAgrément CLng(Cmb_NumDemande), gstrVALIDATION_ANNULÉE
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
        modDAO.GérerTransaction daoANNULATION: Exit Sub
    End If

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
        Case mesSÉLECTION_MANQUANTE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, cboAdoptions(cboDÉCISION), "de la date de la commission ayant décidé le refus ou le retrait")
        Case mesSAISIE_MANQUANTE:       pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, pstrMessages(0))
        Case Else:                      pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub

Private Sub cmdAdoptions_Click(Index As Integer)
Const pkstrNomProcédure As String = "cmdAdoptions_Click"
Dim pitmRecoursContentieuxSélectionné As ListItem, pitmDernierRecoursContentieux As ListItem
Dim pstrMessages(0) As String
Dim plngTraitement As CommunTraitements
Dim plngNuméroRecours As Long
On Error GoTo Err_cmdAdoptions_Click

'   Référencer le recours contentieux sélectionné
    Set pitmRecoursContentieuxSélectionné = lvwAdoptions(lvwCONTENTIEUX).SelectedItem
    
    Select Case Index
    
'       Créer le recours contentieux
        Case cmdCRÉER
        
'           Si un recours contentieux est sélectionné
            If Not pitmRecoursContentieuxSélectionné Is Nothing Then
            
'               Référencer le dernier recours contentieux déposé
                Set pitmDernierRecoursContentieux = lvwAdoptions(lvwCONTENTIEUX).ListItems(1)
    
'               Refuser la création si le plus haut recours contentieux possible a été déposé
                If Not CLng(pitmDernierRecoursContentieux.Text) < gclsRecoursContentieux.lngPlusHautRecours Then
                    Err.Raise mesMESSAGE_008
                    Exit Sub
                End If
            
'               Refuser la création si le dernier recours contentieux déposé n'a pas conduit à une décision
                If pitmDernierRecoursContentieux.SubItems(hdrDÉCISION) = gclsDélibérations(CStr(AdoptionsDélibérations.adpINDÉTERMINÉ)).strLibellé Then
                    Err.Raise mesMESSAGE_010
                    Exit Sub
                End If
            
                plngNuméroRecours = CLng(pitmDernierRecoursContentieux.Text) + 1
                
'           Sinon
            Else
            
                plngNuméroRecours = 1
                
            End If
            
            plngTraitement = comCRÉATION
        
'       Modifier le recours contentieux
        Case cmdMODIFIER
        
'           Refuser la mise à jour si aucun recours n'est sélectionné
            If pitmRecoursContentieuxSélectionné Is Nothing Then
                pstrMessages(0) = "Mise à jour": Err.Raise mesSÉLECTION_MANQUANTE
                Exit Sub
            End If
        
            plngTraitement = comMODIFICATION: plngNuméroRecours = CLng(pitmRecoursContentieuxSélectionné.Text)
        
'       Supprimer le recours contentieux
        Case cmdSUPPRIMER
        
'           Refuser la suppression si aucun recours n'est sélectionné
            If pitmRecoursContentieuxSélectionné Is Nothing Then
                pstrMessages(0) = "Suppression": Err.Raise mesSÉLECTION_MANQUANTE
                Exit Sub
            End If
        
'           Refuser la suppression s'il ne s'agit pas du dernier recours
            If Not pitmRecoursContentieuxSélectionné.Index = 1 Then
                Err.Raise mesMESSAGE_009
                Exit Sub
            End If
        
            plngTraitement = comSUPPRESSION: plngNuméroRecours = CLng(pitmRecoursContentieuxSélectionné.Text)
        
    End Select
    
'   Ouvrir la fenêtre permettant le traitement du recours contentieux
    With mfrmRecoursContentieux
        .Initialiser plngTraitement, Cmb_NumDemande, CDate(cboAdoptions(cboDÉCISION)), plngNuméroRecours
        .Show vbModal
    End With
    
    Exit Sub
    
Err_cmdAdoptions_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:             pintTraitement = mesINTERROMPRE
        Case mesSÉLECTION_MANQUANTE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, "d'un recours contentieux", pstrMessages(0) & " refusée")
        Case mesMESSAGE_008:            pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
        Case mesMESSAGE_009:            pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
        Case mesMESSAGE_010:            pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, "dernier")
        Case Else:                      pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub DTP_RecAdm_Change()
Const pkstrNomProcédure As String = "DTP_RecAdm_Change"
On Error GoTo Err_DTP_RecAdm_Change

'   Activer/Désactiver les contrôles se rapportant au recours administratif
    Implémenter "<activation du recours administratif>"
    
    Exit Sub
    
Err_DTP_RecAdm_Change:
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

    Set mclsDécisions = Nothing
    Set mrstRecoursAdministratif = Nothing: Set mrstRecoursContentieux = Nothing
    Set mfrmRecoursContentieux = Nothing: Set mfrm_TS = Nothing
    
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
        Case cmdPSYCHOLOGUE:        mvarPsychologueNuméro = pvarIntervenantNuméro:          Txt_psy = pstrIntervenantNom
        Case cmdASSISTANTE_SOCIALE: mvarAssistanteSocialeNuméro = pvarIntervenantNuméro:    Txt_as = pstrIntervenantNom
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

Private Sub mfrmRecoursContentieux_Validation(plngTraitement As CommunTraitements, plngNuméroRecours As Long)
Const pkstrNomProcédure As String = "mfrmRecoursContentieux_Validation"
Dim pitmRecours As ListItem
On Error GoTo Err_mfrmRecoursContentieux_Validation

'   Implémenter l'historique des recours contentieux
    Implémenter "Historique des recours contentieux"
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub
    
'   S'il ne s'agit pas d'une suppression
    If Not (plngTraitement = comSUPPRESSION Or plngNuméroRecours = daoENREGISTREMENT_SUPPRIMÉ) Then
'       Référencer le recours sur lequel a porté l'opération
        Set pitmRecours = lvwAdoptions(lvwCONTENTIEUX).ListItems("*" & CStr(plngNuméroRecours))
        pitmRecours.Selected = True
    End If
    
    Exit Sub
    
Err_mfrmRecoursContentieux_Validation:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Implémenter les objets
'   pstrListeObjets:        liste des objets à implémenter au format '<objet#1>[|<objet#2>[...|<objet#n]]'
'   pstrListeParamètres:    liste des paramètres au format '<paramètre#1 de l'objet#1[¤paramètre#2 de l'objet#1[...¤paramètre#n de l'objet#1]]>[|<paramètre#1 de l'objet#2[¤paramètre#2 de l'objet#2[...¤paramètre#n de l'objet#2]]>[...|<paramètre#1 de l'objet#n[¤paramètre#2 de l'objet#n[...¤paramètre#n de l'objet#n]]>]]'

Private Sub Implémenter(pstrListeObjets As String, Optional pstrListeParamètres As String = "")
Const pkstrNomProcédure As String = "Implémenter"
Dim pitmRecoursContentieux As ListItem
Dim pstrListesObjets() As String, pstrListesParamètres() As String, pstrParamètres() As String
Dim pintItération As Integer
Dim pbooVerrouillé As Boolean, pbooRefus As Boolean
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
        
'           Date de la commission ayant décidé le refus ou le retrait
            Case "Décision du"
        
'               Réinitialiser la liste des décisions
                Set mclsDécisions = Nothing: Set mclsDécisions = New clsDécisions
                cboAdoptions(cboDÉCISION).Clear
                modCommun.ImplémenterListe cboAdoptions(cboDÉCISION), "", pvarNuméro:=0
                cboAdoptions(cboDÉCISION).ListIndex = 0
    
'           Historique des recours contentieux
            Case "Historique des recours contentieux"
        
                lvwAdoptions(lvwCONTENTIEUX).ListItems.Clear
                
'               Si les recours contentieux n'ont pas encore été lus
                If mrstRecoursContentieux Is Nothing Then
'                   Les lire
                    Set mrstRecoursContentieux = modDAO.rstSQLExtraction(daoRECOURS_CONTENTIEUX, True, gstrCONSULTATION_LIMITÉE, dbOpenSnapshot, CLng(Cmb_NumDemande), CDate(cboAdoptions(cboDÉCISION)))
'               Sinon
                Else
'                   Rafraîchir le jeu d'enregistrements
                    modDAO.ActualiserRecordset mrstRecoursContentieux, True, gstrCONSULTATION_LIMITÉE
                End If
                If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
                    modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
                    Exit Sub
                End If
                
                With mrstRecoursContentieux
'                   Pour chacun
                    While Not .EOF
'                       Créer un item sans image
                        Set pitmRecoursContentieux = lvwAdoptions(lvwCONTENTIEUX).ListItems.Add(, "*" & CStr(![TYRE_NUM]), ![TYRE_NUM])
                        modCommun.ImplémenterSubItems pitmRecoursContentieux, modCommun.varValeurFormatéeSiNull(![RECO_DDREC], vbDate, ""), gclsRecoursContentieux.clsRecoursContentieux(![TYRE_NUM]).strLibellé, IIf(IsNull(![RECO_DECI]), gclsDélibérations(CStr(AdoptionsDélibérations.adpINDÉTERMINÉ)).strLibellé, IIf(![RECO_DECI] = gstrREFUS, gclsDélibérations(CStr(adpREFUS)).strLibellé, gclsDélibérations(CStr(adpACCORD)).strLibellé)), modCommun.varValeurFormatéeSiNull(![RECO_DDECI], vbDate, ""), modCommun.varValeurFormatéeSiNull(![RECO_dtmAgrementDelivrance], vbDate, ""), modCommun.varValeurFormatéeSiNull(![RECO_REFUS], vbString)
                        .MoveNext
                    Wend
                    txtAdoptions(txtCONTENTIEUX) = CStr(.RecordCount) & mstrNOMBRE_RECOURS
                End With
        
'           Gestion de l'activation des contrôles se rapportant au recours administratif
            Case "<activation du recours administratif>"
        
                pbooVerrouillé = IsNull(DTP_RecAdm)
                If Not Cmb_DecifAdm.ListIndex = comINDÉFINI Then pbooRefus = (Cmb_DecifAdm.ItemData(Cmb_DecifAdm.ListIndex) = adpREFUS)
                
                modCommun.DéfinirVerrouillageContrôle DTP_dcomAdm, pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle DTP_ddeciAdm, pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle DTP_dpredecAdm, pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle Cmb_DeciAdm, pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle Cmb_DecifAdm, pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle Txt_RefusAdm, pbooVerrouillé Or Not pbooRefus
                modCommun.DéfinirVerrouillageContrôle Cmd_TS(cmdASSISTANTE_SOCIALE), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle Cmd_TS(cmdPSYCHOLOGUE), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle DTP_ddemasAdm, pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle DTP_recasAdm, pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle DTP_dprrecasAdm, pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle DTP_ddempsyAdm, pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle DTP_recpsyAdm, pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle DTP_dprrecpsyAdm, pbooVerrouillé
                
'           Dates se rapportant au recours administratif
            Case "<dates du recours administratif>"
        
'               Désélectionner les dates
                DTP_RecAdm = "": DTP_dcomAdm = "": DTP_ddeciAdm = "": DTP_dpredecAdm = "": DTP_ddemasAdm = "": DTP_recasAdm = "": DTP_dprrecasAdm = "": DTP_ddempsyAdm = "": DTP_recpsyAdm = "": DTP_dprrecpsyAdm = ""
                
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
Dim pstrMessages(2) As String, pstrNuméroDossier As String
On Error GoTo Err_Initialiser
    
    mbooInitialisationAvecDemandeDéfinie = Not IsMissing(pvarNuméroDemande)
    
'   Centrer la feuille
    modAdoptions.CentrerFeuille Me
        
'   Implémenter les images
    cmd_Quitter.Picture = Menu.picImage(adp32x32, adpANNULER)
    cmd_Valider.Picture = Menu.picImage(adp32x32, adpVALIDER)
    Cmd_TS(cmdPSYCHOLOGUE).Picture = Menu.picImage(adp16x16, adpINTERVENANTS)
    Cmd_TS(cmdASSISTANTE_SOCIALE).Picture = Menu.picImage(adp16x16, adpINTERVENANTS)
    cmdAdoptions(cmdCRÉER).Picture = Menu.picImage(adp16x16, adpCRÉER)
    cmdAdoptions(cmdMODIFIER).Picture = Menu.picImage(adp16x16, adpMODIFIER)
    cmdAdoptions(cmdSUPPRIMER).Picture = Menu.picImage(adp16x16, adpSUPPRIMER)
    
'   Désactiver certains contrôles
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtDEMANDEUR), True
    modCommun.DéfinirVerrouillageContrôle DTP_DateDemande, True
    modCommun.DéfinirVerrouillageContrôle Txt_as, True
    modCommun.DéfinirVerrouillageContrôle Txt_psy, True
    
'   Vérifier le paramétrage des recours contentieux
    With gclsRecoursContentieux
        If .booParamétrageIncomplet Then
            pstrMessages(0) = .strTribunalAdministratif: pstrMessages(1) = .strCourAdministrativeAppel: pstrMessages(2) = .strConseilÉtat
            Err.Raise mesPARAMÉTRAGE_CONTENTIEUX_INCOMPLET
            modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
            Exit Sub
        End If
    End With

    Set mfrmRecoursContentieux = New frmRecoursContentieux: Set mfrm_TS = New frm_TS
    mvarPsychologueNuméro = Null: mvarAssistanteSocialeNuméro = Null
    
'   Implémenter l'avis administratif
    modAdoptions.ImplémenterListe Cmb_DeciAdm, "Délibérations", "", Nothing, False
    Cmb_DeciAdm.ListIndex = modCommun.varRecherche(Cmb_DeciAdm, False, AdoptionsDélibérations.adpINDÉTERMINÉ)
'   Implémenter la décision finale administrative
    modAdoptions.ImplémenterListe Cmb_DecifAdm, "Délibérations", "", Nothing, False
    Cmb_DecifAdm.ListIndex = modCommun.varRecherche(Cmb_DecifAdm, False, AdoptionsDélibérations.adpINDÉTERMINÉ)
    
'   Implémenter la date de dépôt de la demande, la date de la commission ayant décidé le refus ou le retrait et les dates se rapportant au recours administratif
    Implémenter "Déposée le,Décision du,<dates du recours administratif>", "Faux" & gstrSÉPARATEUR_1 & "" & gstrSÉPARATEUR_1 & ""
    
'   Si la demande est définie
    If mbooInitialisationAvecDemandeDéfinie Then
        pstrNuméroDossier = modAdoptions.txtNuméroDossier(CLng(pvarNuméroDemande), cmd_Valider)
        If pstrNuméroDossier = "" Then Exit Sub
        modCommun.DéfinirVerrouillageContrôle Cmb_NumDossiers, True, True
        modCommun.DéfinirVerrouillageContrôle Cmb_NumDemande, True, True
        modCommun.ImplémenterListe Cmb_NumDossiers, pstrNuméroDossier
        modCommun.ImplémenterListe Cmb_NumDemande, CStr(pvarNuméroDemande)
        Cmb_NumDossiers.ListIndex = modCommun.varRecherche(Cmb_NumDossiers, True, pstrNuméroDossier)
        Cmb_NumDemande.ListIndex = modCommun.varRecherche(Cmb_NumDemande, True, CStr(pvarNuméroDemande))
'   Sinon
    Else
        modAdoptions.ImplémenterListe Cmb_NumDossiers, "Dossiers", gstrCONSULTATION_LIMITÉE, cmd_Valider
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

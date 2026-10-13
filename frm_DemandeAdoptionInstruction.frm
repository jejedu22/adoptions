VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_DemandeAdoptionInstruction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instruction de la demande"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11040
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
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
      Height          =   2295
      Index           =   1
      Left            =   120
      TabIndex        =   46
      Top             =   8400
      Width           =   8655
      Begin VB.CommandButton cmdAdoptions 
         Height          =   375
         Index           =   5
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Supprimer le rapport de l'assistante sociale"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAdoptions 
         Height          =   375
         Index           =   4
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Modifier le rapport de l'assistante sociale"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAdoptions 
         Height          =   375
         Index           =   3
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Créer le rapport d'une assistante sociale"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin MSComctlLib.ListView lvwAdoptions 
         Height          =   975
         Index           =   1
         Left            =   240
         TabIndex        =   52
         ToolTipText     =   "Historique des rapports"
         Top             =   975
         Width           =   8175
         _ExtentX        =   14420
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "<Assistante sociale>"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Nom"
            Object.Width           =   4445
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Demandé le"
            Object.Width           =   1843
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "<Numéro de rapport>"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Reçu le"
            Object.Width           =   1843
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Prévu le"
            Object.Width           =   1843
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "<Cadre d'intervention>"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Cadre"
            Object.Width           =   4322
         EndProperty
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Historique des rapports"
         Height          =   195
         Index           =   5
         Left            =   6840
         TabIndex        =   48
         Top             =   0
         Width           =   1620
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Assistantes sociales"
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
         TabIndex        =   47
         Top             =   0
         Width           =   1725
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
      Height          =   2295
      Index           =   0
      Left            =   120
      TabIndex        =   39
      Top             =   6000
      Width           =   8655
      Begin VB.CommandButton cmdAdoptions 
         Height          =   375
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Créer le rapport d'un psychologue"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAdoptions 
         Height          =   375
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Modifier le rapport du psychologue"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAdoptions 
         Height          =   375
         Index           =   2
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Supprimer le rapport du psychologue"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin MSComctlLib.ListView lvwAdoptions 
         Height          =   975
         Index           =   0
         Left            =   240
         TabIndex        =   45
         ToolTipText     =   "Historique des rapports"
         Top             =   975
         Width           =   8175
         _ExtentX        =   14420
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "<Psychologue>"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Nom"
            Object.Width           =   4445
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Demandé le"
            Object.Width           =   1843
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "<Numéro de rapport>"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Reçu le"
            Object.Width           =   1843
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Prévu le"
            Object.Width           =   1843
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "<Cadre d'intervention>"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Cadre"
            Object.Width           =   4322
         EndProperty
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Historique des rapports"
         Height          =   195
         Index           =   3
         Left            =   6840
         TabIndex        =   41
         Top             =   0
         Width           =   1620
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Psychologues"
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
         TabIndex        =   40
         Top             =   0
         Width           =   1185
      End
   End
   Begin VB.Frame Fra_Pièces 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   8655
      Begin MSComCtl2.DTPicker DTP_RecepDos 
         Height          =   285
         Left            =   5070
         TabIndex        =   16
         ToolTipText     =   "Date de réception du dossier"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_RessourcesREC 
         Height          =   285
         Left            =   5070
         TabIndex        =   38
         ToolTipText     =   "Date de réception des justificatifs"
         Top             =   2820
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_RessourcesDEM 
         Height          =   285
         Left            =   3255
         TabIndex        =   37
         ToolTipText     =   "Date de demande des justificatifs"
         Top             =   2820
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_NeuroPsyREC 
         Height          =   285
         Left            =   5070
         TabIndex        =   35
         ToolTipText     =   "Date de réception de l'attestation"
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_NeuroPsyDEM 
         Height          =   285
         Left            =   3255
         TabIndex        =   34
         ToolTipText     =   "Date de demande de l'attestation"
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_PriseSangREC 
         Height          =   285
         Left            =   5070
         TabIndex        =   32
         ToolTipText     =   "Date de réception des prises de sang"
         Top             =   2220
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_PriseSangDEM 
         Height          =   285
         Left            =   3255
         TabIndex        =   31
         ToolTipText     =   "Date de demande des prises de sang"
         Top             =   2220
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_MedecinGenREC 
         Height          =   285
         Left            =   5070
         TabIndex        =   29
         ToolTipText     =   "Date de réception des certificats"
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_MedecinGenDEM 
         Height          =   285
         Left            =   3255
         TabIndex        =   28
         ToolTipText     =   "Date de demande des certificats"
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_CasiersJudREC 
         Height          =   285
         Left            =   5070
         TabIndex        =   26
         ToolTipText     =   "Date de réception des extraits"
         Top             =   1620
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_CasiersJudDEM 
         Height          =   285
         Left            =   3255
         TabIndex        =   25
         ToolTipText     =   "Date de demande des extraits"
         Top             =   1620
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_ActesNaissREC 
         Height          =   285
         Left            =   5070
         TabIndex        =   20
         ToolTipText     =   "Date de réception de la copie"
         Top             =   1020
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_ActesNaissDEM 
         Height          =   285
         Left            =   3255
         TabIndex        =   19
         ToolTipText     =   "Date de demande de la copie"
         Top             =   1020
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_PrevueDos 
         Height          =   285
         Left            =   6885
         TabIndex        =   17
         ToolTipText     =   "Date prévue pour la réception du dossier"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_DateDemande 
         Height          =   285
         Left            =   3255
         TabIndex        =   15
         ToolTipText     =   "Date de la demande du dossier"
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
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_FFrec 
         Height          =   285
         Left            =   5070
         TabIndex        =   23
         ToolTipText     =   "Date de réception de la fiche"
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker DTP_FFdem 
         Height          =   285
         Left            =   3255
         TabIndex        =   22
         ToolTipText     =   "Date de demande de la fiche"
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59244547
         CurrentDate     =   2958465
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Pièces"
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
         TabIndex        =   54
         Top             =   0
         Width           =   585
      End
      Begin VB.Label Lbl_fichefam 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fiche familiale"
         Height          =   195
         Left            =   2025
         TabIndex        =   21
         Top             =   1365
         Width           =   990
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00FF0000&
         X1              =   240
         X2              =   8400
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Lbl_AttestationNP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Attestation du neuro-psychiatre"
         Height          =   195
         Left            =   825
         TabIndex        =   33
         Top             =   2565
         Width           =   2190
      End
      Begin VB.Label Lbl_Ressources 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Justificatifs de ressources"
         Height          =   195
         Left            =   1215
         TabIndex        =   36
         Top             =   2865
         Width           =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   0
         X1              =   3135
         X2              =   3135
         Y1              =   360
         Y2              =   3180
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         Index           =   0
         X1              =   4950
         X2              =   4950
         Y1              =   360
         Y2              =   3180
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF0000&
         Index           =   0
         X1              =   6765
         X2              =   6765
         Y1              =   360
         Y2              =   3180
      End
      Begin VB.Label Lbl_ActeNaissance 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Copie intégrale des actes de naissance"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1065
         Width           =   2775
      End
      Begin VB.Label Lbl_Demandées 
         Alignment       =   2  'Center
         Caption         =   "Demande le"
         Height          =   255
         Left            =   3255
         TabIndex        =   11
         Top             =   345
         Width           =   1575
      End
      Begin VB.Label Lbl_Recues 
         Alignment       =   2  'Center
         Caption         =   "Réception le"
         Height          =   255
         Left            =   5070
         TabIndex        =   12
         Top             =   345
         Width           =   1575
      End
      Begin VB.Label Lbl_CasiersJudiciaires 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Extraits du casier judiciaire"
         Height          =   195
         Left            =   1155
         TabIndex        =   24
         Top             =   1665
         Width           =   1860
      End
      Begin VB.Label Lbl_Certificat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Certificats du médecin généraliste"
         Height          =   195
         Left            =   645
         TabIndex        =   27
         Top             =   1965
         Width           =   2370
      End
      Begin VB.Label Lbl_PrisesDeSang 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Prises de sang"
         Height          =   195
         Left            =   1980
         TabIndex        =   30
         Top             =   2265
         Width           =   1035
      End
      Begin VB.Label Lbl_prévues 
         Alignment       =   2  'Center
         Caption         =   "Prévue le"
         Height          =   255
         Left            =   6885
         TabIndex        =   13
         Top             =   345
         Width           =   1575
      End
      Begin VB.Label Lbl_Dossier 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dossier"
         Height          =   195
         Left            =   2490
         TabIndex        =   14
         Top             =   765
         Width           =   525
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
      Width           =   8655
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
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   7
         ToolTipText     =   "Demandeur"
         Top             =   360
         Width           =   4560
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
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Demandeur"
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   6
         Top             =   405
         Width           =   825
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
      Width           =   8655
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   7080
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
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Valider"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar stb_DemandesInstruction 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   53
      Top             =   10785
      Width           =   8880
      _ExtentX        =   15663
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
            TextSave        =   "01/10/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "14:58"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9437
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_DemandeAdoptionInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
                          
Private WithEvents mfrmIntervenant As frmIntervenant
Attribute mfrmIntervenant.VB_VarHelpID = -1

Private Enum IndexColumnHeader              'Index des contrôles ColumnHeader
    hdrINTERVENANT_NUMÉRO = 0
    hdrINTERVENANT_NOM
    hdrDEMANDÉ
    hdrRAPPORT_NUMÉRO
    hdrPRÉVU
    hdrREÇU
    hdrCADRE_NUMÉRO
    hdrCADRE_LIBELLÉ
End Enum
Private Enum IndexCommandButton             'Index des contrôles CommandButton
    cmdPSYCHOLOGUE_CRÉER = 0
    cmdPSYCHOLOGUE_MODIFIER
    cmdPSYCHOLOGUE_SUPPRIMER
    cmdASSISTANTE_SOCIALE_CRÉER
    cmdASSISTANTE_SOCIALE_MODIFIER
    cmdASSISTANTE_SOCIALE_SUPPRIMER
End Enum
Private Enum IndexFrame                     'Index des contrôles Frame
    fraPSYCHOLOGUES = 0
    fraASSISTANTES_SOCIALES
End Enum
Private Enum IndexListView                  'Index des contrôles ListView
    lvwPSYCHOLOGUES = 0
    lvwASSISTANTES_SOCIALES
End Enum
Private Enum IndexTextBox                   'Index des contrôles TextBox
    txtDEMANDEUR = 0
End Enum
Private Enum Divers                         'Constantes diverses
    NOMBRE_LIGNES_VISIBLES = 3
End Enum

Private mrstDemande As Recordset, mrstPsychologues As Recordset, mrstAssistantesSociales As Recordset
Private mlngIntervenant As IndexCommandButton
Private msngLargeurCadre As Single
Private mbooInitialisationAvecDemandeDéfinie As Boolean, mbooGestionIntervenant As Boolean, mbooValidationRefusée As Boolean

Private Sub Cmb_NumDemande_Click()
Const pkstrNomProcédure As String = "Cmb_NumDemande_Click"
Dim prstIntervenants As Recordset
Dim pfldDemande As Field, pfldNumero As Field, pfldPrevu As Field, pfldRecu As Field
Dim plngListView As IndexListView
Dim pbooInformationsNonConcordantes As Boolean, pbooExistePsychologueInitial As Boolean, pbooExistePsychologueEnCharge As Boolean, pbooExisteAssistanteSocialeInitiale As Boolean, pbooExisteAssistanteSocialeEnCharge As Boolean
On Error GoTo Err_Cmb_NumDemande_Click

'   Implémenter le titre
    Me.Caption = "Instruction de la demande n° " & Cmb_NumDemande.Text & " / dossier n° " + Cmb_NumDossier
    
'   Positionner une liste de DTPicker à la date du jour
    modCommun.PositionnerDateJour DTP_RecepDos, DTP_PrevueDos, DTP_ActesNaissDEM, DTP_ActesNaissREC, DTP_FFdem, DTP_FFrec, DTP_CasiersJudDEM, DTP_CasiersJudREC, DTP_MedecinGenDEM, DTP_MedecinGenREC, DTP_PriseSangDEM, DTP_PriseSangREC, DTP_NeuroPsyDEM, DTP_NeuroPsyREC, DTP_RessourcesDEM, DTP_RessourcesREC
                
'   Activer les frames
    Implémenter "<désactivation des frames>", "Faux"
    
'   Lire la demande
    Set mrstDemande = modDAO.rstSQLExtraction(daoEXTRACTION_025, True, gstrCONSULTATION_LIMITÉE & gstrSÉPARATEUR_1 & gstrDEMANDE_SUPPRIMÉE, CLng(Cmb_NumDemande))
'   Déterminer si la validation est autorisée
    Implémenter "<interdiction de valider>", CStr((Not modDAO.lngDernièreErreurDAO() = daoAUCUNE))
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then Exit Sub
    
'   Implémenter les historiques des intervenants
    Implémenter "Historiques des intervenants", "Vrai" & gstrSÉPARATEUR_2 & "Vrai" & gstrSÉPARATEUR_2 & "Vrai"
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub
    
    Do
    
        pbooInformationsNonConcordantes = False
        pbooExistePsychologueInitial = False: pbooExistePsychologueEnCharge = False: pbooExisteAssistanteSocialeInitiale = False: pbooExisteAssistanteSocialeEnCharge = False
        
'       Pour chaque historique
        For plngListView = lvwPSYCHOLOGUES To lvwASSISTANTES_SOCIALES
        
            Select Case plngListView
                Case lvwPSYCHOLOGUES:           Set prstIntervenants = mrstPsychologues
                Case lvwASSISTANTES_SOCIALES:   Set prstIntervenants = mrstAssistantesSociales
            End Select
            
            With prstIntervenants
            
                If Not .RecordCount = 0 Then .MoveFirst
            
'               Pour chaque intervenant
                While Not .EOF
                
'                   S'il s'agit de l'intervenant en charge
                    If .AbsolutePosition = 0 Then
                    
                        Select Case plngListView
'                           Historique des psychologues
                            Case lvwPSYCHOLOGUES:           pbooExistePsychologueEnCharge = True:       Set pfldDemande = mrstDemande![DEMA_DDEMRP]:            Set pfldNumero = mrstDemande![TRSO_NUMPSY]:         Set pfldPrevu = mrstDemande![DEMA_DPRRP]:           Set pfldRecu = mrstDemande![DEMA_DRECRP]
'                           Historique des assistantes sociales
                            Case lvwASSISTANTES_SOCIALES:   pbooExisteAssistanteSocialeEnCharge = True: Set pfldDemande = mrstDemande![DEMA_DDEMRA]:            Set pfldNumero = mrstDemande![TRSO_NUMAS]:          Set pfldPrevu = mrstDemande![DEMA_DPRRA]:           Set pfldRecu = mrstDemande![DEMA_DRECRA]
                        End Select
                    
'                       Vérifier que les informations de la demande relatives à l'intervenant en charge correspondent à celles de l'historique
                        pbooInformationsNonConcordantes = pbooInformationsNonConcordantes Or Not (![u01_dtmHistoIntervenantDemande] = varValeurFormatéeSiNull(pfldDemande, vbDate) And ![TRSO_NUM] = varValeurFormatéeSiNull(pfldNumero, vbInteger, comINDÉFINI) And varValeurFormatéeSiNull(![u01_dtmHistoIntervenantPrevu], vbDate) = varValeurFormatéeSiNull(pfldPrevu, vbDate) And varValeurFormatéeSiNull(![u01_dtmHistoIntervenantRecu], vbDate) = varValeurFormatéeSiNull(pfldRecu, vbDate))
                        
                    End If
                    
'                   Si l'historique se rapporte à la demande de rapport initiale
                    If ![u01_intHistoIntervenantCadre] = AdoptionsCadresInterventions.adpDEMANDE_INITIALE Then
                    
                        Select Case plngListView
'                           Historique des psychologues
                            Case lvwPSYCHOLOGUES:           pbooExistePsychologueInitial = True:        Set pfldDemande = mrstDemande![DEMA_DDEMRP_Initial]:  Set pfldNumero = mrstDemande![TRSO_NUMPSY_Initial]: Set pfldPrevu = mrstDemande![DEMA_DPRRP_Initial]:   Set pfldRecu = mrstDemande![DEMA_DRECRP_Initial]
'                           Historique des assistantes sociales
                            Case lvwASSISTANTES_SOCIALES:   pbooExisteAssistanteSocialeInitiale = True: Set pfldDemande = mrstDemande![DEMA_DDEMRA_Initial]:       Set pfldNumero = mrstDemande![TRSO_NUMAS_Initial]:  Set pfldPrevu = mrstDemande![DEMA_DPRRA_Initial]:   Set pfldRecu = mrstDemande![DEMA_DRECRA_Initial]
                        End Select
                    
'                       Vérifier que les informations de la demande relatives à l'intervenant initial correspondent à celles de l'historique
                        pbooInformationsNonConcordantes = pbooInformationsNonConcordantes Or Not (![u01_dtmHistoIntervenantDemande] = varValeurFormatéeSiNull(pfldDemande, vbDate) And ![TRSO_NUM] = varValeurFormatéeSiNull(pfldNumero, vbInteger, comINDÉFINI) And varValeurFormatéeSiNull(![u01_dtmHistoIntervenantPrevu], vbDate) = varValeurFormatéeSiNull(pfldPrevu, vbDate) And varValeurFormatéeSiNull(![u01_dtmHistoIntervenantRecu], vbDate) = varValeurFormatéeSiNull(pfldRecu, vbDate))
                        
                    End If
                
                    .MoveNext
                    
                Wend
                
            End With
    
        Next plngListView
                    
'       Vérifier l'existence d'un historique pour chaque intervenant présent sur la demande
        pbooInformationsNonConcordantes = pbooInformationsNonConcordantes Or Not (Not pbooExistePsychologueEnCharge = IsNull(mrstDemande![TRSO_NUMPSY]))
        pbooInformationsNonConcordantes = pbooInformationsNonConcordantes Or Not (Not pbooExisteAssistanteSocialeEnCharge = IsNull(mrstDemande![TRSO_NUMAS]))
        pbooInformationsNonConcordantes = pbooInformationsNonConcordantes Or Not (Not pbooExistePsychologueInitial = IsNull(mrstDemande![TRSO_NUMPSY_Initial]))
        pbooInformationsNonConcordantes = pbooInformationsNonConcordantes Or Not (Not pbooExisteAssistanteSocialeInitiale = IsNull(mrstDemande![TRSO_NUMAS_Initial]))
        
'       Si les informations relatives aux intervenants présents sur la demande ne correspondent pas à celles de l'historique
        If pbooInformationsNonConcordantes Then
        
'           Si l'utilisateur décide leur mise à jour
            If modMessages.lngRéponseUtilisateur(mesQUESTION_007) = vbYes Then
            
'               Ouvrir une transaction
                modDAO.GérerTransaction daoOUVERTURE
                
'               Actualiser les données relatives aux intervenants présents sur la demande
                modAdoptions.ActualiserDonnéesIntervenants CLng(Cmb_NumDemande), gstrCONSULTATION_LIMITÉE
                If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
                    modDAO.GérerTransaction daoANNULATION
'                   Interdire la validation
                    Implémenter "<interdiction de valider>", "Vrai"
                    Exit Sub
                End If
            
'               Valider la transaction
                modDAO.GérerTransaction daoVALIDATION
            
'               Rafraîchir la demande
                modDAO.ActualiserRecordset mrstDemande, True, gstrCONSULTATION_LIMITÉE
                If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
'                   Interdire la validation
                    Implémenter "<interdiction de valider>", "Vrai"
                    Exit Sub
                End If
                If modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT Then
                    Err.Raise mesENREGISTREMENT_SUPPRIMÉ
'                   Interdire la validation
                    Implémenter "<interdiction de valider>", "Vrai"
                    Exit Sub
                End If
            
'           Sinon
            Else
            
                pbooInformationsNonConcordantes = False
                
            End If
            
        End If
                
    Loop While pbooInformationsNonConcordantes
    
'   Positionner les contrôles
    With mrstDemande
        DTP_DateDemande = modCommun.varValeurFormatéeSiNull(![DEMA_DATE], vbDate, Null)
        DTP_RecepDos = modCommun.varValeurFormatéeSiNull(![DEMA_DREDOS], vbDate, Null)
        DTP_PrevueDos = modCommun.varValeurFormatéeSiNull(![DEMA_DPRDOS], vbDate, Null)
        DTP_ActesNaissDEM = modCommun.varValeurFormatéeSiNull(![DEMA_DDEMAN], vbDate, ![DEMA_DATE]):    If IsNull(![DEMA_DDEMAN]) Then DTP_ActesNaissDEM = ""
        DTP_ActesNaissREC = modCommun.varValeurFormatéeSiNull(![DEMA_DRECAN], vbDate, Null)
        DTP_FFdem = modCommun.varValeurFormatéeSiNull(![DEMA_DDEMFF], vbDate, ![DEMA_DATE]):            If IsNull(![DEMA_DDEMFF]) Then DTP_FFdem = ""
        DTP_FFrec = modCommun.varValeurFormatéeSiNull(![DEMA_DRECFF], vbDate, Null)
        DTP_CasiersJudDEM = modCommun.varValeurFormatéeSiNull(![DEMA_DDEMCJ], vbDate, ![DEMA_DATE]):    If IsNull(![DEMA_DDEMCJ]) Then DTP_CasiersJudDEM = ""
        DTP_CasiersJudREC = modCommun.varValeurFormatéeSiNull(![DEMA_DRECCJ], vbDate, Null)
        DTP_MedecinGenDEM = modCommun.varValeurFormatéeSiNull(![DEMA_DDEMCM], vbDate, ![DEMA_DATE]):    If IsNull(![DEMA_DDEMCM]) Then DTP_MedecinGenDEM = ""
        DTP_MedecinGenREC = modCommun.varValeurFormatéeSiNull(![DEMA_DRECCM], vbDate, Null)
        DTP_PriseSangDEM = modCommun.varValeurFormatéeSiNull(![DEMA_DDEMPS], vbDate, ![DEMA_DATE]):     If IsNull(![DEMA_DDEMPS]) Then DTP_PriseSangDEM = ""
        DTP_PriseSangREC = modCommun.varValeurFormatéeSiNull(![DEMA_DRECPS], vbDate, Null)
        DTP_NeuroPsyDEM = modCommun.varValeurFormatéeSiNull(![DEMA_DDEMNP], vbDate, ![DEMA_DATE]):      If IsNull(![DEMA_DDEMNP]) Then DTP_NeuroPsyDEM = ""
        DTP_NeuroPsyREC = modCommun.varValeurFormatéeSiNull(![DEMA_DRECNP], vbDate, Null)
        DTP_RessourcesDEM = modCommun.varValeurFormatéeSiNull(![DEMA_DDEMJR], vbDate, ![DEMA_DATE]):    If IsNull(![DEMA_DDEMJR]) Then DTP_RessourcesDEM = ""
        DTP_RessourcesREC = modCommun.varValeurFormatéeSiNull(![DEMA_DRECJR], vbDate, Null)
    End With
    
'   Activer/Désactiver l'ensemble des contrôles se rapportant au psychologue et à l'assistante sociale selon que la date de réception du dossier est saisie ou non
    DTP_RecepDos_Change
        
    Exit Sub
    
Err_Cmb_NumDemande_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                 pintTraitement = mesINTERROMPRE
        Case mesENREGISTREMENT_SUPPRIMÉ:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, "Consultation", "La demande")
        Case Else:                          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
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
Dim pvarDEMA_DREF As Variant
Dim pctlFocus As Control
Dim pdtpDemandé As DTPicker, pdtpReçu As DTPicker, pdtpPrévu As DTPicker
Dim pstrMessages(1) As String, pstrChaîne As String
Dim pintItération As Integer
Dim pbooPiècesMultiples As Boolean
On Error GoTo Err_cmd_Valider_Click

    mbooValidationRefusée = False
    
'   Forcer si nécessaire l'actualisation de la valeur d'un contrôle DTPicker
    modCommun.GérerActualisationDTPicker Me, cmd_Valider
    
'   Refuser la validation si le numéro de la demande n'est pas sélectionné
    If Cmb_NumDemande = "" Then
        pstrMessages(0) = "du numéro de la demande": Err.Raise mesSÉLECTION_MANQUANTE
        mbooValidationRefusée = True: Exit Sub
    End If
    
'   Déterminer la date de référence pour le calcul de l'échéancier
    pvarDEMA_DREF = modCommun.varValeurFormatéeSiNull(mrstDemande![DEMA_DREF], vbDate, DTP_RecepDos)
    
'   Si la date prévue pour la réception du dossier est désactivée
    If IsNull(DTP_PrevueDos) Then
'       La recalculer
        DTP_PrevueDos = DateAdd("m", modDAO.intParamètreÉchéancier(daoDEMANDE__RÉCEPTION_DOSSIER), mrstDemande![DEMA_DATE])
    End If
    
'   Refuser la validation si la date de réception du dossier est antérieure à la date de sa demande
    If DTP_DateDemande > modCommun.varValeurFormatéeSiNull(DTP_RecepDos, vbDate) Then
        Set pctlFocus = DTP_RecepDos: pstrMessages(0) = "de réception du dossier" & vbCrLf: pstrMessages(1) = "de sa demande"
        Err.Raise mesDATE_ANTÉRIEURE
        mbooValidationRefusée = True: Exit Sub
    End If

'   Refuser la validation si la date prévue pour la réception du dossier est antérieure à la date de sa demande
    If DTP_DateDemande > DTP_PrevueDos Then
        Set pctlFocus = DTP_PrevueDos: pstrMessages(0) = "prévue pour la réception du dossier" & vbCrLf: pstrMessages(1) = "de sa demande"
        Err.Raise mesDATE_ANTÉRIEURE
        mbooValidationRefusée = True: Exit Sub
    End If

'   Pour chaque pièce
    For pintItération = 1 To 7
    
        Select Case pintItération
            Case 1: Set pdtpDemandé = DTP_ActesNaissDEM:    Set pdtpReçu = DTP_ActesNaissREC:   pstrChaîne = "de la copie intégrale des actes de naissance":    pbooPiècesMultiples = False
            Case 2: Set pdtpDemandé = DTP_FFdem:            Set pdtpReçu = DTP_FFrec:           pstrChaîne = "de la fiche familiale":                           pbooPiècesMultiples = False
            Case 3: Set pdtpDemandé = DTP_CasiersJudDEM:    Set pdtpReçu = DTP_CasiersJudREC:   pstrChaîne = "des extraits du casier judiciaire":               pbooPiècesMultiples = True
            Case 4: Set pdtpDemandé = DTP_MedecinGenDEM:    Set pdtpReçu = DTP_MedecinGenREC:   pstrChaîne = "des certificats du médecin généraliste":          pbooPiècesMultiples = True
            Case 5: Set pdtpDemandé = DTP_PriseSangDEM:     Set pdtpReçu = DTP_PriseSangREC:    pstrChaîne = "des prises de sang":                              pbooPiècesMultiples = True
            Case 6: Set pdtpDemandé = DTP_NeuroPsyDEM:      Set pdtpReçu = DTP_NeuroPsyREC:     pstrChaîne = "de l'attestation du neuro-psychiatre":            pbooPiècesMultiples = False
            Case 7: Set pdtpDemandé = DTP_RessourcesDEM:    Set pdtpReçu = DTP_RessourcesREC:   pstrChaîne = "des justificatifs de ressources":                 pbooPiècesMultiples = True
        End Select
        
'       Refuser la validation si les dates relatives à la pièce sont incohérentes
        If booSaisieIncohérentePièces(pdtpDemandé, pdtpReçu, pstrChaîne, pbooPiècesMultiples) Then mbooValidationRefusée = True: Exit Sub

    Next pintItération
    
'   Refuser la validation si la date de réception n'est pas définie mais que des intervenants sont saisis
    If IsNull(DTP_RecepDos) And Not (lvwAdoptions(lvwPSYCHOLOGUES).ListItems.Count = 0 And lvwAdoptions(lvwASSISTANTES_SOCIALES).ListItems.Count = 0) Then
        Set pctlFocus = DTP_RecepDos: pstrMessages(0) = "la saisie d'intervenants"
        Err.Raise mesDATE_RÉCEPTION_ABSENTE
        mbooValidationRefusée = True: Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    With mrstDemande
    
        .Edit
        
'       Mettre à jour la demande
        ![DEMA_DREDOS] = DTP_RecepDos
        ![DEMA_DPRDOS] = DTP_PrevueDos
        ![DEMA_DDEMAN] = DTP_ActesNaissDEM
        ![DEMA_DRECAN] = DTP_ActesNaissREC
        ![DEMA_DDEMFF] = DTP_FFdem
        ![DEMA_DRECFF] = DTP_FFrec
        ![DEMA_DDEMCJ] = DTP_CasiersJudDEM
        ![DEMA_DRECCJ] = DTP_CasiersJudREC
        ![DEMA_DDEMCM] = DTP_MedecinGenDEM
        ![DEMA_DRECCM] = DTP_MedecinGenREC
        ![DEMA_DDEMPS] = DTP_PriseSangDEM
        ![DEMA_DRECPS] = DTP_PriseSangREC
        ![DEMA_DDEMNP] = DTP_NeuroPsyDEM
        ![DEMA_DRECNP] = DTP_NeuroPsyREC
        ![DEMA_DDEMJR] = DTP_RessourcesDEM
        ![DEMA_DRECJR] = DTP_RessourcesREC
        ![DEMA_DREF] = pvarDEMA_DREF
                
'       Valider la mise à jour
        modDAO.ValiderOpérationSurRecordset mrstDemande
        
        Select Case modDAO.lngDernièreErreurDAO()
        
'           L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
            Case daoENREGISTREMENT_VERROUILLÉ
                Screen.MousePointer = vbDefault: mbooValidationRefusée = True: Exit Sub
                
'           L'opération a été refusée en raison de la suppression par un autre utilisateur des données originales
            Case dbRecordDeleted
                Err.Raise mesENREGISTREMENT_SUPPRIMÉ
'               Interdire la validation
                Implémenter "<interdiction de valider>", "Vrai"
                Screen.MousePointer = vbDefault: mbooValidationRefusée = True: Exit Sub
                
'           L'opération a été refusée en raison de la modification par un autre utilisateur des données originales
            Case dbRecordModified
            
'           L'opération a généré un conflit inattendu
            Case Is <> daoAUCUNE
                Err.Raise mesCONFLIT_ACCÈS
                
        End Select
    
    End With
    
    Screen.MousePointer = vbDefault
    
    If Not mbooGestionIntervenant Then Unload Me
    
    Exit Sub
  
Err_cmd_Valider_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                 pintTraitement = mesINTERROMPRE
        Case mesCONFLIT_ACCÈS:              pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, modDAO.lngDernièreErreurDAO())
        Case mesSÉLECTION_MANQUANTE:        pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Cmb_NumDemande, pstrMessages(0))
        Case mesDATE_ANTÉRIEURE:            pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, pstrMessages(0), pstrMessages(1))
        Case mesENREGISTREMENT_SUPPRIMÉ:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, "Consultation", "La demande")
        Case mesDATE_RÉCEPTION_ABSENTE:     pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, pstrMessages(0))
        Case Else:                          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub

Private Sub cmdAdoptions_Click(Index As Integer)
Const pkstrNomProcédure As String = "cmdAdoptions_Click"
Dim pclsCadreIntervention As clsCodeLibellé
Dim pitmIntervenantSélectionné As ListItem
Dim pstrMessages(2) As String, pstrChaînes(1) As String
Dim plngListView As IndexListView, plngCadreIntervention As AdoptionsCadresInterventions, plngTraitement As CommunTraitements
Dim pbooPsychologue As Boolean
On Error GoTo Err_cmdAdoptions_Click

    Select Case Index
'       Psychologues
        Case cmdPSYCHOLOGUE_CRÉER, cmdPSYCHOLOGUE_MODIFIER, cmdPSYCHOLOGUE_SUPPRIMER:                       plngListView = lvwPSYCHOLOGUES:         pbooPsychologue = True
'       Assistantes sociales
        Case cmdASSISTANTE_SOCIALE_CRÉER, cmdASSISTANTE_SOCIALE_MODIFIER, cmdASSISTANTE_SOCIALE_SUPPRIMER:  plngListView = lvwASSISTANTES_SOCIALES: pbooPsychologue = False
    End Select
        
'   Référencer le recours contentieux sélectionné
    Set pitmIntervenantSélectionné = lvwAdoptions(plngListView).SelectedItem
    
    Select Case Index
    
'       Créer l'intervenant
        Case cmdPSYCHOLOGUE_CRÉER, cmdASSISTANTE_SOCIALE_CRÉER
    
'           Si la date de réception du sossier ou la date de référence pour le calcul de l'échéancier n'est pas définie
            If IsNull(mrstDemande![DEMA_DREDOS]) Or IsNull(mrstDemande![DEMA_DREF]) Then
'               Proposer la validation de l'instruction
                Select Case modMessages.lngRéponseUtilisateur(mesQUESTION_008)
'                   L'utilisateur décide la validation
                    Case vbYes
                        mbooGestionIntervenant = True
                        cmd_Valider_Click
                        mbooGestionIntervenant = False
                        If mbooValidationRefusée Then Exit Sub
'                   L'utilisateur décide l'abandon du traitement
                    Case vbNo
                        Exit Sub
                End Select
            End If
            
            plngTraitement = comCRÉATION
            
'           Ouvrir la fenêtre permettant le traitement de l'intervenant
            With mfrmIntervenant
                .Initialiser plngTraitement, Cmb_NumDemande, pbooPsychologue
                .Show vbModal
            End With
            
'       Modifier ou supprimer l'intervenant
        Case cmdPSYCHOLOGUE_MODIFIER, cmdPSYCHOLOGUE_SUPPRIMER, cmdASSISTANTE_SOCIALE_MODIFIER, cmdASSISTANTE_SOCIALE_SUPPRIMER
    
            Select Case Index
'               Modifier l'intervenant
                Case cmdPSYCHOLOGUE_MODIFIER, cmdASSISTANTE_SOCIALE_MODIFIER:   plngTraitement = comMODIFICATION:   pstrMessages(0) = "Mise à jour":    pstrMessages(1) = "mises à jour"
'               Supprimer l'intervenant
                Case cmdPSYCHOLOGUE_SUPPRIMER, cmdASSISTANTE_SOCIALE_SUPPRIMER: plngTraitement = comSUPPRESSION:    pstrMessages(0) = "Suppression":    pstrMessages(1) = "supprimées"
            End Select
            
'           Refuser la mise à jour si aucun intervenant n'est sélectionné
            If pitmIntervenantSélectionné Is Nothing Then
                pstrMessages(2) = IIf(pbooPsychologue, "d'un psychologue", "d'une assistante sociale")
                Err.Raise mesSÉLECTION_MANQUANTE
                Exit Sub
            End If
    
            plngCadreIntervention = pitmIntervenantSélectionné.SubItems(hdrCADRE_NUMÉRO)
            
'           Refuser la mise à jour si le cadre d'intervention définit un recours administratif ou une surveillance de placement
            If plngCadreIntervention = AdoptionsCadresInterventions.adpRECOURS_ADMINISTRATIF Or plngCadreIntervention = adpSURVEILLANCE_PLACEMENT Then
                Set pclsCadreIntervention = gclsCadresInterventions(CStr(plngCadreIntervention))
                pstrMessages(0) = pclsCadreIntervention.strUn & LCase(pclsCadreIntervention.strLibellé)
                Err.Raise mesMESSAGE_011
                Exit Sub
            End If
        
'           S'il s'agit de mettre à jour l'intervenant
            If plngTraitement = comMODIFICATION Then
'               Si les dates de réception du dossier ou de référence pour le calcul de l'échéancier ne sont pas définies
                If IsNull(mrstDemande![DEMA_DREDOS]) Or IsNull(mrstDemande![DEMA_DREF]) Then
'                   Proposer la validation de l'instruction
                    Select Case modMessages.lngRéponseUtilisateur(mesQUESTION_008)
'                       L'utilisateur décide la validation
                        Case vbYes
                            mbooGestionIntervenant = True
                            cmd_Valider_Click
                            mbooGestionIntervenant = False
                            If mbooValidationRefusée Then Exit Sub
'                       L'utilisateur décide l'abandon du traitement
                        Case vbNo
                            Exit Sub
                    End Select
                End If
            End If
            
'           Ouvrir la fenêtre permettant le traitement de l'intervenant
            With mfrmIntervenant
                .Initialiser plngTraitement, Cmb_NumDemande, pbooPsychologue, CDate(pitmIntervenantSélectionné.SubItems(hdrDEMANDÉ)), CInt(pitmIntervenantSélectionné.SubItems(hdrRAPPORT_NUMÉRO))
                .Show vbModal
            End With
    
    End Select
        
    Exit Sub
    
Err_cmdAdoptions_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:             pintTraitement = mesINTERROMPRE
        Case mesSÉLECTION_MANQUANTE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, pstrMessages(2), pstrMessages(0) & " refusée")
        Case mesMESSAGE_011:            pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, pstrMessages(0), pstrMessages(1))
        Case Else:                      pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub DTP_RecepDos_Change()
Const pkstrNomProcédure As String = "DTP_RecepDos_Change"
Dim pbooVerrouillé As Boolean
On Error GoTo Err_DTP_RecepDos_Change

'   Activer/Désactiver les contrôles se rapportant se rapportant aux intervenants
    Implémenter "<activation des intervenants>"
    
    Exit Sub
    
Err_DTP_RecepDos_Change:
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

    Set mrstDemande = Nothing: Set mrstPsychologues = Nothing: Set mrstAssistantesSociales = Nothing
    Set mfrmIntervenant = Nothing
    
    Exit Sub
    
Err_Form_Unload:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub

Private Sub mfrmIntervenant_Validation(plngTraitement As CommunTraitements, pbooPsychologue As Boolean, pdtmRapportDemandé As Date, pintRapportNuméro As Integer)
Const pkstrNomProcédure As String = "mfrmIntervenant_Validation"
Dim pitmIntervenant As ListItem
On Error GoTo Err_mfrmIntervenant_Validation

'   Implémenter les historiques des intervenants
    Implémenter "Historiques des intervenants", "Faux" & gstrSÉPARATEUR_2 & CStr(pbooPsychologue) & gstrSÉPARATEUR_2 & CStr(Not pbooPsychologue)
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub
    
'   S'il ne s'agit pas d'une suppression
    If Not (plngTraitement = comSUPPRESSION Or CLng(pintRapportNuméro) = daoENREGISTREMENT_SUPPRIMÉ) Then
'       Référencer l'intervenant sur lequel a porté l'opération
        Set pitmIntervenant = lvwAdoptions(IIf(pbooPsychologue, lvwPSYCHOLOGUES, lvwASSISTANTES_SOCIALES)).ListItems(CStr(pdtmRapportDemandé) & gstrSÉPARATEUR_1 & CStr(pintRapportNuméro))
        pitmIntervenant.Selected = True
    End If
    
    Exit Sub
    
Err_mfrmIntervenant_Validation:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Retourner un booléen précisant si la saisie des dates relatives aux pièces est incohérente
'   pdtpDemandé:            DTPicker permettant la saisie de la date de demande de la pièce
'   pdtpReçu:               DTPicker permettant la saisie de la date de réception de la pièce
'   pstrChaîne:             chaîne utilisée dans la construction du message en cas de saisie incohérente (voir pvarArgument(0) de mesDATE_ANTÉRIEURE) [Exemple: "de la fiche familiale"]
'   pbooPiècesMultiples:    booléen précisant si la pièce est constituée de plusieurs pièces

Private Function booSaisieIncohérentePièces(pdtpDemandé As DTPicker, pdtpReçu As DTPicker, pstrChaîne As String, pbooPiècesMultiples As Boolean) As Boolean
Const pkstrNomProcédure As String = "booSaisieIncohérentePièces"
Dim pdtpFocus As DTPicker
Dim pstrMessages(1) As String
On Error GoTo Err_booSaisieIncohérentePièces

'   Refuser la validation si la demande de la pièce est antérieure à la demande du dossier
    If DTP_DateDemande > modCommun.varValeurFormatéeSiNull(pdtpDemandé, vbDate) Then
        Set pdtpFocus = pdtpDemandé: pstrMessages(0) = "de demande " & pstrChaîne & vbCrLf: pstrMessages(1) = "de demande du dossier"
        Err.Raise mesDATE_ANTÉRIEURE
        booSaisieIncohérentePièces = True
        Exit Function
    End If

'   Refuser la validation si la date de réception de la pièce est antérieure à la demande du dossier
    If DTP_DateDemande > modCommun.varValeurFormatéeSiNull(pdtpReçu, vbDate) Then
        Set pdtpFocus = pdtpReçu: pstrMessages(0) = "de réception " & pstrChaîne & vbCrLf: pstrMessages(1) = "de demande du dossier"
        Err.Raise mesDATE_ANTÉRIEURE
        booSaisieIncohérentePièces = True
        Exit Function
    End If

'   Refuser la validation si la date de réception de la pièce est antérieure à la date de sa demande
    If modCommun.varValeurFormatéeSiNull(pdtpDemandé, vbDate, #12:00:00 AM#) > modCommun.varValeurFormatéeSiNull(pdtpReçu, vbDate) Then
        Set pdtpFocus = pdtpReçu: pstrMessages(0) = "de réception " & pstrChaîne & vbCrLf: pstrMessages(1) = "de " & IIf(pbooPiècesMultiples, "leur", "sa") & " demande"
        Err.Raise mesDATE_ANTÉRIEURE
        booSaisieIncohérentePièces = True
        Exit Function
    End If
    
    Exit Function
    
Err_booSaisieIncohérentePièces:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:         pintTraitement = mesINTERROMPRE
        Case mesDATE_ANTÉRIEURE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pdtpFocus, pstrMessages(0), pstrMessages(1))
        Case Else:                  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Function


'Implémenter les objets
'   pstrListeObjets:        liste des objets à implémenter au format '<objet#1>[|<objet#2>[...|<objet#n]]'
'   pstrListeParamètres:    liste des paramètres au format '<paramètre#1 de l'objet#1[¤paramètre#2 de l'objet#1[...¤paramètre#n de l'objet#1]]>[|<paramètre#1 de l'objet#2[¤paramètre#2 de l'objet#2[...¤paramètre#n de l'objet#2]]>[...|<paramètre#1 de l'objet#n[¤paramètre#2 de l'objet#n[...¤paramètre#n de l'objet#n]]>]]'

Private Sub Implémenter(pstrListeObjets As String, Optional pstrListeParamètres As String = "")
Const pkstrNomProcédure As String = "Implémenter"
Dim pitmIntervenant As ListItem
Dim prstIntervenants As Recordset
Dim pstrListesObjets() As String, pstrListesParamètres() As String, pstrParamètres() As String
Dim plngListView As IndexListView
Dim pintItération As Integer
Dim pbooVerrouillé As Boolean, pbooLire As Boolean, pbooImplémenterHistorique As Boolean, pbooTraiterPsychologues As Boolean
On Error GoTo Err_Implémenter
    
    pstrListesObjets() = Split(pstrListeObjets, ","): pstrListesParamètres() = Split(pstrListeParamètres, gstrSÉPARATEUR_1)
    
'   Pour chaque objet dont on souhaite l'implémentation
    For pintItération = 0 To UBound(pstrListesObjets)
    
        Select Case pstrListesObjets(pintItération)
        
'           Gestion de l'activation des contrôles se rapportant aux intervenants
            Case "<activation des intervenants>"
        
                pbooVerrouillé = IsNull(DTP_RecepDos)
    
                modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdPSYCHOLOGUE_CRÉER), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdPSYCHOLOGUE_MODIFIER), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdPSYCHOLOGUE_SUPPRIMER), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle lvwAdoptions(lvwPSYCHOLOGUES), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdASSISTANTE_SOCIALE_CRÉER), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdASSISTANTE_SOCIALE_MODIFIER), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdASSISTANTE_SOCIALE_SUPPRIMER), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle lvwAdoptions(lvwASSISTANTES_SOCIALES), pbooVerrouillé
                
'           Gestion de l'activation des frames
            Case "<désactivation des frames>"
        
                pstrParamètres() = Split(pstrListesParamètres(pintItération), gstrSÉPARATEUR_2)
                pbooVerrouillé = CBool(pstrParamètres(0))
                
                modCommun.DéfinirVerrouillageContrôle Fra_Pièces, pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle fraAdoptions(fraPSYCHOLOGUES), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle fraAdoptions(fraASSISTANTES_SOCIALES), pbooVerrouillé
                
'           Historiques des intervenants
            Case "Historiques des intervenants"
        
                pstrParamètres() = Split(pstrListesParamètres(pintItération), gstrSÉPARATEUR_2)
                pbooLire = CBool(pstrParamètres(0))
                
                For plngListView = lvwPSYCHOLOGUES To lvwASSISTANTES_SOCIALES
                
                    Select Case plngListView
                        Case lvwPSYCHOLOGUES:           Set prstIntervenants = mrstPsychologues:        pbooImplémenterHistorique = CBool(pstrParamètres(1)):   pbooTraiterPsychologues = True
                        Case lvwASSISTANTES_SOCIALES:   Set prstIntervenants = mrstAssistantesSociales: pbooImplémenterHistorique = CBool(pstrParamètres(2)):   pbooTraiterPsychologues = False
                    End Select
                    
'                   Si l'historique doit être implémenté
                    If pbooImplémenterHistorique Then
                    
                        lvwAdoptions(plngListView).ListItems.Clear
                        
'                       Si les intervenants n'ont pas encore été lus ou s'ils doivent l'être
                        If prstIntervenants Is Nothing Or pbooLire Then
'                           Les lire
                            Set prstIntervenants = modDAO.rstSQLExtraction(daoPSYCHOLOGUES__OU__ASSISTANTES_SOCIALES, True, gstrCONSULTATION_LIMITÉE, dbOpenSnapshot, CLng(Cmb_NumDemande), pbooTraiterPsychologues)
                            If plngListView = lvwPSYCHOLOGUES Then Set mrstPsychologues = prstIntervenants Else Set mrstAssistantesSociales = prstIntervenants
'                       Sinon
                        Else
'                           Rafraîchir le jeu d'enregistrements
                            modDAO.ActualiserRecordset prstIntervenants, True, gstrCONSULTATION_LIMITÉE
                        End If
                        If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
'                           Interdire la validation
                            Implémenter "<interdiction de valider>", "Vrai"
                            Exit Sub
                        End If
                        
                        With prstIntervenants
'                           Pour chacun
                            While Not .EOF
'                               Créer un item sans image
                                Set pitmIntervenant = lvwAdoptions(plngListView).ListItems.Add(, CStr(![u01_dtmHistoIntervenantDemande]) & gstrSÉPARATEUR_1 & CStr(![u01_intHistoIntervenantNumero]), ![TRSO_NUM])
                                modCommun.ImplémenterSubItems pitmIntervenant, modCommun.varValeurFormatéeSiNull(![TRSO_NOM], vbString), ![u01_dtmHistoIntervenantDemande], ![u01_intHistoIntervenantNumero], modCommun.varValeurFormatéeSiNull(![u01_dtmHistoIntervenantRecu], vbDate, ""), modCommun.varValeurFormatéeSiNull(![u01_dtmHistoIntervenantPrevu], vbDate, ""), ![u01_intHistoIntervenantCadre], gclsCadresInterventions(CStr(![u01_intHistoIntervenantCadre])).strLibellé
                                .MoveNext
                            Wend
                        End With
                
                        lvwAdoptions(plngListView).ColumnHeaders(hdrCADRE_LIBELLÉ + 1).Width = msngLargeurCadre - IIf(lvwAdoptions(plngListView).ListItems.Count > NOMBRE_LIGNES_VISIBLES, comLARGEUR_ASCENCEUR, 0)
        
                    End If
                
                Next plngListView
                
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
    cmdAdoptions(cmdPSYCHOLOGUE_CRÉER).Picture = Menu.picImage(adp16x16, adpCRÉER)
    cmdAdoptions(cmdPSYCHOLOGUE_MODIFIER).Picture = Menu.picImage(adp16x16, adpMODIFIER)
    cmdAdoptions(cmdPSYCHOLOGUE_SUPPRIMER).Picture = Menu.picImage(adp16x16, adpSUPPRIMER)
    cmdAdoptions(cmdASSISTANTE_SOCIALE_CRÉER).Picture = Menu.picImage(adp16x16, adpCRÉER)
    cmdAdoptions(cmdASSISTANTE_SOCIALE_MODIFIER).Picture = Menu.picImage(adp16x16, adpMODIFIER)
    cmdAdoptions(cmdASSISTANTE_SOCIALE_SUPPRIMER).Picture = Menu.picImage(adp16x16, adpSUPPRIMER)

'   Désactiver certains contrôles
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtDEMANDEUR), True
    modCommun.DéfinirVerrouillageContrôle DTP_DateDemande, True
    
    Set mfrmIntervenant = New frmIntervenant
    msngLargeurCadre = lvwAdoptions(lvwPSYCHOLOGUES).ColumnHeaders(hdrCADRE_LIBELLÉ + 1).Width

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

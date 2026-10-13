VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_PlacementEnfant 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Placement d'un enfant"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   ControlBox      =   0   'False
   Icon            =   "frm_PlacementEnfant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fra_enfant 
      Caption         =   "Placement d'un enfant"
      Height          =   2175
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   7335
      Begin VB.TextBox Txt_PaysNaissance 
         Height          =   285
         Left            =   4560
         MaxLength       =   20
         TabIndex        =   5
         ToolTipText     =   "Pays de naissance de l'enfant placé"
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox Txt_LieuNaissance 
         Height          =   285
         Left            =   4560
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "Lieu de naissance de l'enfant placé"
         Top             =   240
         Width           =   2535
      End
      Begin VB.ComboBox Cmb_Sexe 
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Sexe de l'enfant placé"
         Top             =   960
         Width           =   495
      End
      Begin MSComCtl2.DTPicker DTP_DateNaissance 
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         ToolTipText     =   "Date de naissance de l'enfant placé"
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59113475
         CurrentDate     =   36307
      End
      Begin VB.TextBox Txt_DepNaissance 
         Height          =   285
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   3
         ToolTipText     =   "Département de naissance"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Txt_Prenom 
         Height          =   285
         Left            =   720
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "Prénom de l'enfant placé"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Txt_Nom 
         Height          =   285
         Left            =   720
         MaxLength       =   20
         TabIndex        =   0
         ToolTipText     =   "Nom de l'enfant placé"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   7200
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line3 
         X1              =   7200
         X2              =   7200
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   120
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Label Lbl_Notes 
         Caption         =   "Notes : les champs ""Nom"", ""Prénom"" et ""Date de naissance"" sont obligatoires."
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   6855
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7200
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Lbl_Sexe 
         Caption         =   "Sexe"
         Height          =   255
         Left            =   3120
         TabIndex        =   18
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Lbl_PaysNaissance 
         Caption         =   "Pays de naissance"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Lbl_LieuNaissance 
         Caption         =   "Lieu de naissance"
         Height          =   255
         Left            =   3120
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Lbl_DepNaissance 
         Caption         =   "Département de naissance"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Lbl_DateNaissance 
         Caption         =   "Date de naissance"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Lbl_Prenom 
         Caption         =   "Prénom"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Lbl_Nom 
         Caption         =   "Nom"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   6720
         Picture         =   "frm_PlacementEnfant.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   6000
         Picture         =   "frm_PlacementEnfant.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_DemandeEnfant 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2790
      Width           =   7530
      _ExtentX        =   13282
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
            TextSave        =   "01/04/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11:51"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7056
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_PlacementEnfant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private QuitterCroix As Boolean

Private Sub Cmb_Sexe_Gotfocus()
    stb_DemandeEnfant.SimpleText = "Choisissez le sexe de l'enfant"
End Sub

Private Sub cmd_Quitter_Click()
    QuitterCroix = False
    Unload Me
End Sub

Private Sub cmd_Valider_Click()
    Dim requete As String
    
    QuitterCroix = False
    On Error GoTo Rupture_Integrite
    Workspaces(0).BeginTrans
    
    Select Case stb_DemandeEnfant.Panels(4).Text
        Case "CREATION"
            requete = "INSERT INTO ENPL(ENPL_NOM,ENPL_PREN,ENPL_DNAIS,PLAC_NUM,ENPL_DPNAIS,ENPL_LNAISS,ENPL_PAYS,ENPL_SEXE)"
            requete = requete + "VALUES("
            requete = requete + Iifch(Txt_Nom) + ","
            requete = requete + Iifch(Txt_Prenom) + ","
            requete = requete + Iifch(DTP_DateNaissance) + ","
            requete = requete + Iifent(frm_PlacementGestion.Txt_NumeroPlacement) + ","
            requete = requete + Iifent(Txt_DepNaissance) + ","
            requete = requete + Iifch(Txt_LieuNaissance) + ","
            requete = requete + Iifch(Txt_PaysNaissance) + ","
            requete = requete + Iifch(Cmb_Sexe.Text) + ")"
            
        Case "MODIFICATION"
            requete = "UPDATE ENPL SET"
            requete = requete + " ENPL_DNAIS=" + Iifch(DTP_DateNaissance)
            requete = requete + ",ENPL_DPNAIS=" + Iifent(Txt_DepNaissance)
            requete = requete + ",ENPL_LNAISS=" + Iifch(Txt_LieuNaissance)
            requete = requete + ",ENPL_PAYS=" + Iifch(Txt_PaysNaissance)
            requete = requete + ",ENPL_SEXE=" + Iifch(Cmb_Sexe.Text)
            requete = requete + " WHERE PLAC_NUM = " + Iifent(frm_PlacementGestion.Txt_NumeroPlacement)
            requete = requete + " AND ENPL_NOM = " + Iifch(Txt_Nom)
            requete = requete + " AND ENPL_PREN = " + Iifch(Txt_Prenom)
            
        Case "SUPPRESSION"
            requete = "DELETE * FROM ENPL WHERE PLAC_NUM = " + Iifent(frm_PlacementGestion.Txt_NumeroPlacement)
            requete = requete + " AND ENPL_NOM = " + Iifch(Txt_Nom)
            requete = requete + " AND ENPL_PREN = " + Iifch(Txt_Prenom)
    End Select
    Workspaces(0).Databases(0).Execute requete, dbFailOnError
    Workspaces(0).CommitTrans
    Unload Me

    Exit Sub
    
Rupture_Integrite:
'Affiche le code d'erreur approprié
    Select Case Err
        Case 13
            MsgBox "Valeur incompatible avec le type de colonne", vbOKOnly + vbExclamation
      
        Case 3022
            MsgBox strDoublon + ", création impossible", vbOKOnly + vbExclamation
      
        Case 3200
            MsgBox Reference + ", suppression impossible", vbOKOnly + vbExclamation
      
        Case 3075
            MsgBox "L'un des champs contient une cote, mise à jour impossible", vbOKOnly + vbExclamation
      
        Case 3315
            MsgBox "Un champ obligatoire n'est pas renseigné, mise à jour impossible", vbOKOnly + vbExclamation
      
        Case Else
            MsgBox "Erreur" + CStr(Err) + " " + Err.Description + " survenue, mise à jour impossible", vbOKOnly + vbExclamation
    End Select
    Workspaces(0).Rollback
    MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub DTP_DateNaissance_GotFocus()
    stb_DemandeEnfant.SimpleText = "Entrez la date de naissance de l'enfant"
End Sub

Private Sub Form_Load()
    QuitterCroix = True
    ListeSexe Cmb_Sexe
    Cmb_Sexe.ListIndex = 0
End Sub

Public Sub ListeSexe(Boite As ComboBox)
'Permet la sélection du sexe de l'enfant gràce à la combo
    Boite.AddItem ("M")
    Boite.AddItem ("F")
End Sub


Private Sub Txt_DepNaissance_Gotfocus()
    stb_DemandeEnfant.SimpleText = "Entrez le département de naissance de l'enfant"
End Sub

Private Sub Txt_LieuNaissance_Gotfocus()
    stb_DemandeEnfant.SimpleText = "Entrez le lieu de naissance de l'enfant"
End Sub

Private Sub Txt_Nom_Gotfocus()
    stb_DemandeEnfant.SimpleText = "Entrez le nom de l'enfant"
End Sub

Private Sub Txt_PaysNaissance_Gotfocus()
    stb_DemandeEnfant.SimpleText = "Entrez le pays de naissance de l'enfant"
End Sub

Private Sub Txt_Prenom_Gotfocus()
    stb_DemandeEnfant.SimpleText = "Entrez le prénom de l'enfant"
End Sub

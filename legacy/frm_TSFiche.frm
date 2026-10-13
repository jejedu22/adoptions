VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_TSFiche 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Travailleur social"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_TSFiche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fra_ts 
      Height          =   4095
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   9255
      Begin VB.CheckBox chk_actif 
         Caption         =   "En activité"
         Height          =   255
         Left            =   3480
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox Cmb_Titr 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         ToolTipText     =   "Titre de l'intervenant"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Data Dat_Titre 
         Caption         =   "Dat_Titre"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data Dat_ReqPerm 
         Caption         =   "Dat_ReqPerm"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox Cmb_Perm 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Lieu de permanance où officie l'intervenant"
         Top             =   3480
         Width           =   4575
      End
      Begin VB.Data Dat_ReqQual 
         Caption         =   "Dat_ReqQual"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox Cmb_Qual 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Qualité de l'intervenant"
         Top             =   3000
         Width           =   4575
      End
      Begin VB.ComboBox Cmb_Inst 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
         Height          =   315
         ItemData        =   "frm_TSFiche.frx":0442
         Left            =   2040
         List            =   "frm_TSFiche.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Institution dans laquelle travaille l'intervenant"
         Top             =   2520
         Width           =   4575
      End
      Begin VB.Data Dat_ReqInst 
         Caption         =   "Dat_ReqInst"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "INST"
         Top             =   2520
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txt_telprof 
         Height          =   285
         Left            =   2040
         MaxLength       =   14
         TabIndex        =   3
         ToolTipText     =   "Téléphone professionnel de l'intervenant"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Txt_code 
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         ToolTipText     =   "Code de l'intervenant"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Txt_nom 
         Height          =   285
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   2
         ToolTipText     =   "Nom de l'intervenant"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   840
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label lbl_referent 
         Caption         =   "Referent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   2160
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   1560
         X2              =   9120
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label lbl_inst 
         Caption         =   "Institution"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lbl_liperm 
         Caption         =   "Lieu de permanence"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label lbl_titr 
         Caption         =   "Titre"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbl_telprof 
         Caption         =   "Tel prof"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lbl_code 
         Caption         =   "Code"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lbl_nom 
         Caption         =   "Nom Prénom"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lbl_qualité 
         Caption         =   "Qualité"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   1695
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   8640
         Picture         =   "frm_TSFiche.frx":0446
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8040
         Picture         =   "frm_TSFiche.frx":0548
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_TrSoFiche 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   4800
      Width           =   9510
      _ExtentX        =   16775
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
            Object.Width           =   10548
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_TSFiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public QuitterCroix As Boolean

Private Sub Cmb_Inst_Gotfocus()
    If (stb_TrSoFiche.Panels(4).Text = "CREATION") Or (stb_TrSoFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TrSoFiche.SimpleText = "Choisissez l'institution de l'intervenant social"
    End If
End Sub

Private Sub Cmb_Perm_Gotfocus()
    If (stb_TrSoFiche.Panels(4).Text = "CREATION") Or (stb_TrSoFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TrSoFiche.SimpleText = "Choisissez la permanence de l'intervenant social"
    End If
End Sub

Private Sub Cmb_Qual_Gotfocus()
    If (stb_TrSoFiche.Panels(4).Text = "CREATION") Or (stb_TrSoFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TrSoFiche.SimpleText = "Choisissez la qualité de l'intervenant social"
    End If
End Sub

Private Sub Cmb_Serv_Gotfocus()
    If (stb_TrSoFiche.Panels(4).Text = "CREATION") Or (stb_TrSoFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TrSoFiche.SimpleText = "Choisissez le service de l'intervenant social"
    End If
End Sub

Private Sub Cmb_Titr_Gotfocus()
    If (stb_TrSoFiche.Panels(4).Text = "CREATION") Or (stb_TrSoFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TrSoFiche.SimpleText = "Choisissez le titre de l'intervenant social"
    End If
End Sub

Private Sub cmd_Quitter_Click()
    QuitterCroix = False
    Unload frm_TSFiche
End Sub

Private Sub cmd_Valider_Click()
    Dim requete As String
    
    QuitterCroix = False
    Dim titre As String
    titre = Cmb_Titr.Text
    On Error GoTo Rupture_Integrite
    Select Case stb_TrSoFiche.Panels(4).Text
        Case "CREATION"
            requete = "INSERT INTO TRSO VALUES ("
            requete = requete + Iifent(Txt_Code) + ","
            requete = requete + CodeEntreCotes(Cmb_Inst.Text) + ","
            requete = requete + CodeEntreCotes(Cmb_Qual.Text) + ","
            requete = requete + CodeEntreCotes(Cmb_Perm.Text) + ","
            requete = requete + Iifch(titre) + ","
            requete = requete + Iifch(Txt_Nom) + ","
            requete = requete + Iifch(txt_telprof) + ","
            requete = requete + CStr(chk_actif.Value) + ")"
            
        Case "MODIFICATION"
            requete = "UPDATE TRSO"
            requete = requete + " SET INST_CODE=" + CodeEntreCotes(Cmb_Inst.Text)
            requete = requete + ",QUAL_CODE=" + CodeEntreCotes(Cmb_Qual.Text)
            requete = requete + ",PERM_CODE=" + CodeEntreCotes(Cmb_Perm.Text)
            requete = requete + ",TRSO_TITR=" + Iifch(Cmb_Titr.Text)
            requete = requete + ",TRSO_NOM=" + Iifch(Txt_Nom)
            requete = requete + ",TRSO_TELPROF=" + Iifch(txt_telprof)
            requete = requete + ",TRSO_ACTIF=" + CStr(chk_actif.Value)
            requete = requete + " WHERE TRSO_NUM=" + Iifent(Txt_Code)
            
        Case "SUPPRESSION"
            requete = "DELETE * FROM TRSO WHERE TRSO_NUM = " + Iifent(Txt_Code)
    End Select
    
    Workspaces(0).Databases(0).Execute requete, dbFailOnError
    Unload Me
    frm_TS.Dat_TS.Refresh
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
    Exit Sub
End Sub

Private Sub Form_Load()

    QuitterCroix = True
    If frm_TS.stb_TS.Panels(4) = "CREATION" Then
        stb_TrSoFiche.Panels(4).Text = "CREATION"
    End If
    If frm_TS.stb_TS.Panels(4) = "MODIFICATION" Then
        stb_TrSoFiche.Panels(4).Text = "MODIFICATION"
    End If
    If frm_TS.stb_TS.Panels(4) = "SUPPRESSION" Then
        stb_TrSoFiche.Panels(4).Text = "SUPPRESSION"
    End If
    
    Dat_ReqInst.DatabaseName = CheminBase
    Dat_ReqQual.DatabaseName = CheminBase
    Dat_ReqPerm.DatabaseName = CheminBase
    
End Sub

Private Sub txt_adr1_gotfocus()
    If (stb_TrSoFiche.Panels(4).Text = "CREATION") Or (stb_TrSoFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TrSoFiche.SimpleText = "Entrez l'adresse de l'intervenant social"
    End If
End Sub

Private Sub txt_adr2_Gotfocus()
    If (stb_TrSoFiche.Panels(4).Text = "CREATION") Or (stb_TrSoFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TrSoFiche.SimpleText = "Entrez l'adresse de l'intervenant social"
    End If
End Sub

Private Sub Txt_code_Gotfocus()
    If (stb_TrSoFiche.Panels(4).Text = "CREATION") Or (stb_TrSoFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TrSoFiche.SimpleText = "Entrez le code de l'intervenant social"
    End If
End Sub

Private Sub txt_cp_Gotfocus()
    If (stb_TrSoFiche.Panels(4).Text = "CREATION") Or (stb_TrSoFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TrSoFiche.SimpleText = "Entrez le code postal de l'intervenant social"
    End If
End Sub

Private Sub Txt_Nom_Gotfocus()
    If (stb_TrSoFiche.Panels(4).Text = "CREATION") Or (stb_TrSoFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TrSoFiche.SimpleText = "Entrez le nom de l'intervenant social"
    End If
End Sub

Private Sub Txt_Tel_Gotfocus()
    If (stb_TrSoFiche.Panels(4).Text = "CREATION") Or (stb_TrSoFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TrSoFiche.SimpleText = "Entrez le téléphone personnel de l'intervenant social"
    End If
End Sub

Private Sub txt_telprof_Gotfocus()
    If (stb_TrSoFiche.Panels(4).Text = "CREATION") Or (stb_TrSoFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TrSoFiche.SimpleText = "Entrez le téléphone professionnel de l'intervenant social"
    End If
End Sub

Private Sub txt_Ville_Gotfocus()
    If (stb_TrSoFiche.Panels(4).Text = "CREATION") Or (stb_TrSoFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TrSoFiche.SimpleText = "Entrez la ville de l'intervenant social"
    End If
End Sub




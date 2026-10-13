VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_TribunauxFiche 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tribunaux Fiche"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_TribunauxFiche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fra_tribunaux 
      Height          =   5295
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   9255
      Begin VB.TextBox txt_corres 
         Height          =   285
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   6
         ToolTipText     =   "Nom du correspondant au sein du tribunal"
         Top             =   3960
         Width           =   3615
      End
      Begin VB.TextBox txt_ville 
         Height          =   285
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   5
         ToolTipText     =   "Ville où se situe le tribunal"
         Top             =   3360
         Width           =   3615
      End
      Begin VB.TextBox txt_cp 
         Height          =   285
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   4
         ToolTipText     =   "Code postal du tribunal"
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txt_adr2 
         Height          =   285
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   3
         ToolTipText     =   "Complément d'adresse"
         Top             =   2160
         Width           =   3615
      End
      Begin VB.TextBox txt_adr1 
         Height          =   285
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   2
         ToolTipText     =   "Adresse du tribunal"
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox Txt_nom 
         Height          =   285
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   1
         ToolTipText     =   "Libellé du tribunal"
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox Txt_code 
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         ToolTipText     =   "Code du tribunal"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbl_Correspondant 
         Caption         =   "Commentaires"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label lbl_ville 
         Caption         =   "Ville"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lbl_cp 
         Caption         =   "Code Postal"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label lbl_adr2 
         Caption         =   "Adresse 2"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lbl_adr1 
         Caption         =   "Adresse 1"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lbl_nomassurma 
         Caption         =   "Libellé"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lbl_code 
         Caption         =   "Code"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   8040
         Picture         =   "frm_TribunauxFiche.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Annuler"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   8640
         Picture         =   "frm_TribunauxFiche.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Valider"
         Top             =   120
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_TribunauxFiche 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   6075
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
Attribute VB_Name = "frm_TribunauxFiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private QuitterCroix As Boolean

Private Sub cmd_Quitter_Click()
    QuitterCroix = False
    Unload Me
End Sub

Private Sub cmd_Valider_Click()
    Dim requete As String
        
    QuitterCroix = False
    On Error GoTo Rupture_Integrite
    Select Case stb_TribunauxFiche.Panels(4).Text
        Case "CREATION"
            requete = "INSERT INTO JUGE(JUGE_NUM,JUGE_NOM,JUGE_ADR1,JUGE_ADR2,JUGE_CPTGI,JUGE_VILLE,JUGE_NOMC)"
            requete = requete + "VALUES ("
            requete = requete + Iifent(Txt_Code) + ","
            requete = requete + Iifch(Txt_Nom) + ","
            requete = requete + Iifch(Txt_adr1) + ","
            requete = requete + Iifch(Txt_adr2) + ","
            requete = requete + Iifch(Txt_CP) + ","
            requete = requete + Iifch(Txt_Ville) + ","
            requete = requete + Iifch(txt_corres) + ");"
        Case "MODIFICATION"
            requete = "UPDATE JUGE"
            requete = requete + " SET JUGE_NOM=" + Iifch(Txt_Nom)
            requete = requete + ",JUGE_ADR1=" + Iifch(Txt_adr1)
            requete = requete + ",JUGE_ADR2=" + Iifch(Txt_adr2)
            requete = requete + ",JUGE_CPTGI=" + Iifch(Txt_CP)
            requete = requete + ",JUGE_VILLE=" + Iifch(Txt_Ville)
            requete = requete + ",JUGE_NOMC=" + Iifch(txt_corres)
            requete = requete + " WHERE JUGE_NUM=" + Iifent(Txt_Code)
        Case "SUPPRESSION"
            requete = "DELETE * FROM JUGE WHERE JUGE_NUM = " + Iifent(Txt_Code)
    End Select
    Workspaces(0).Databases(0).Execute requete, dbFailOnError
    Unload Me
    frm_Tribunaux.Dat_Trib.Refresh
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
    If frm_Tribunaux.stb_ListeTribunaux.Panels(4) = "CREATION" Then
        stb_TribunauxFiche.Panels(4).Text = "CREATION"
    End If
    If frm_Tribunaux.stb_ListeTribunaux.Panels(4) = "MODIFICATION" Then
        stb_TribunauxFiche.Panels(4).Text = "MODIFICATION"
    End If
    If frm_Tribunaux.stb_ListeTribunaux.Panels(4) = "SUPPRESSION" Then
        stb_TribunauxFiche.Panels(4).Text = "SUPPRESSION"
    End If
End Sub

Private Sub txt_adr1_gotfocus()
    If (stb_TribunauxFiche.Panels(4).Text = "CREATION") Or (stb_TribunauxFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TribunauxFiche.SimpleText = "Entrez l'adresse du tribunal"
    End If
End Sub

Private Sub txt_adr2_Gotfocus()
    If (stb_TribunauxFiche.Panels(4).Text = "CREATION") Or (stb_TribunauxFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TribunauxFiche.SimpleText = "Entrez l'adresse du tribunal"
    End If
End Sub

Private Sub Txt_code_Gotfocus()
    If (stb_TribunauxFiche.Panels(4).Text = "CREATION") Or (stb_TribunauxFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TribunauxFiche.SimpleText = "Entrez le code du tribunal"
    End If
End Sub

Private Sub txt_corres_Gotfocus()
    If (stb_TribunauxFiche.Panels(4).Text = "CREATION") Or (stb_TribunauxFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TribunauxFiche.SimpleText = "Entrez le nom du correspondant du tribunal"
    End If
End Sub

Private Sub txt_cp_Gotfocus()
    If (stb_TribunauxFiche.Panels(4).Text = "CREATION") Or (stb_TribunauxFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TribunauxFiche.SimpleText = "Entrez le code postal du tribunal"
    End If
End Sub

Private Sub Txt_Nom_Gotfocus()
    If (stb_TribunauxFiche.Panels(4).Text = "CREATION") Or (stb_TribunauxFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TribunauxFiche.SimpleText = "Entrez le nom du tribunal"
    End If
End Sub

Private Sub txt_Ville_Gotfocus()
    If (stb_TribunauxFiche.Panels(4).Text = "CREATION") Or (stb_TribunauxFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TribunauxFiche.SimpleText = "Entrez la ville du tribunal"
    End If
End Sub




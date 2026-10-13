VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_OeuvreFiche 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oeuvres de placement"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_OeuvreFiche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleMode       =   0  'User
   ScaleWidth      =   9500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   8040
         Picture         =   "frm_OeuvreFiche.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   8640
         Picture         =   "frm_OeuvreFiche.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Frame Fra_oeuvres 
      Height          =   5295
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   9255
      Begin VB.TextBox Txt_adr 
         Height          =   285
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   22
         ToolTipText     =   "Adresse de l'oeuvre"
         Top             =   1560
         Width           =   5775
      End
      Begin VB.TextBox Txt_cp 
         Height          =   285
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   21
         ToolTipText     =   "Code postal"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Txt_ville 
         Height          =   285
         Left            =   3960
         MaxLength       =   20
         TabIndex        =   20
         ToolTipText     =   "Ville de domiciliation"
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Txt_nomres 
         Height          =   285
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   19
         ToolTipText     =   "Nom du responsable de l'association"
         Top             =   3960
         Width           =   4335
      End
      Begin VB.TextBox Txt_rem 
         Height          =   285
         Left            =   2040
         MaxLength       =   60
         TabIndex        =   18
         ToolTipText     =   "Remarques éventuelles"
         Top             =   4560
         Width           =   6975
      End
      Begin VB.TextBox Txt_EMail 
         Height          =   285
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   16
         ToolTipText     =   "Adresse électronique"
         Top             =   3360
         Width           =   4335
      End
      Begin VB.TextBox Txt_tel 
         Height          =   285
         Left            =   2040
         MaxLength       =   14
         TabIndex        =   2
         ToolTipText     =   "Numéro de téléphone"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox Txt_nom 
         Height          =   285
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Nom de l'oeuvre"
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox Txt_code 
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         ToolTipText     =   "Code l'oeuvre de placement"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "E- Mail"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lbl_tel 
         Caption         =   "Téléphone"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lbl_remarques 
         Caption         =   "Remarques"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label lbl_nomresp 
         Caption         =   "Nom du responsable"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label lbl_ville 
         Caption         =   "Ville"
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lbl_cp 
         Caption         =   "Code Postal"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lbl_adresse 
         Caption         =   "Adresse"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lbl_nomasso 
         Caption         =   "Nom de l'association"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lbl_code 
         Caption         =   "Code"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar stb_OeuvreFiche 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   5970
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
Attribute VB_Name = "frm_OeuvreFiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private QuitterCroix As Boolean

Private Sub cmd_Quitter_Click()
    QuitterCroix = False
    Unload frm_OeuvreFiche
End Sub

Private Sub cmd_Valider_Click()

  Dim requete As String
  
  QuitterCroix = False
  On Error GoTo Rupture_Integrite
  
  Select Case stb_OeuvreFiche.Panels(4).Text
      Case "CREATION"
        requete = "INSERT INTO OEPL VALUES("
        requete = requete + Iifent(Txt_Code) + ","
        requete = requete + Iifch(Txt_Nom) + ","
        requete = requete + Iifch(Txt_Adr) + ","
        requete = requete + Iifch(Txt_CP) + ","
        requete = requete + Iifch(Txt_Ville) + ","
        requete = requete + Iifch(Txt_Tel) + ","
        requete = requete + Iifch(Txt_nomres) + ","
        requete = requete + Iifch(Txt_rem) + ","
        requete = requete + Iifch(Txt_EMail) + ")"
        
      Case "MODIFICATION"
          requete = "UPDATE OEPL SET "
          requete = requete + "OEPL_DESIGN=" + Iifch(Txt_Nom)
          requete = requete + ",OEPL_ADR=" + Iifch(Txt_Adr)
          requete = requete + ",OEPL_CP=" + Iifch(Txt_CP)
          requete = requete + ",OEPL_VILLE=" + Iifch(Txt_Ville)
          requete = requete + ",OEPL_TEL=" + Iifch(Txt_Tel)
          requete = requete + ",OEPL_NOMRES=" + Iifch(Txt_nomres)
          requete = requete + ",OEPL_REM=" + Iifch(Txt_rem)
          requete = requete + ",OEPL_EMAIL=" + Iifch(Txt_EMail)
          requete = requete + " WHERE OEPL_CODE=" + Iifent(Txt_Code)
          
      Case "SUPPRESSION"
          requete = "DELETE * FROM OEPL WHERE OEPL_CODE = " + Iifent(Txt_Code)
          
  End Select
  Workspaces(0).Databases(0).Execute requete, dbFailOnError
  Unload Me
  frm_Oeuvres.Dat_oeuvre.Refresh
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
    If frm_Oeuvres.stb_ListeOeuvres.Panels(4) = "CREATION" Then
        stb_OeuvreFiche.Panels(4).Text = "CREATION"
    End If
    If frm_Oeuvres.stb_ListeOeuvres.Panels(4) = "MODIFICATION" Then
        stb_OeuvreFiche.Panels(4).Text = "MODIFICATION"
    End If
    If frm_Oeuvres.stb_ListeOeuvres.Panels(4) = "SUPPRESSION" Then
        stb_OeuvreFiche.Panels(4).Text = "SUPPRESSION"
    End If
End Sub

Private Sub Txt_Adr_Gotfocus()
    If (stb_OeuvreFiche.Panels(4).Text = "CREATION") Or (stb_OeuvreFiche.Panels(4).Text = "MODIFICATION") Then
        stb_OeuvreFiche.SimpleText = "Entrez l'adresse de l'association"
    End If
End Sub

Private Sub Txt_code_Gotfocus()
    If (stb_OeuvreFiche.Panels(4).Text = "CREATION") Or (stb_OeuvreFiche.Panels(4).Text = "MODIFICATION") Then
        stb_OeuvreFiche.SimpleText = "Entrez le numéro de l'association"
    End If
End Sub

Private Sub txt_cp_Gotfocus()
    If (stb_OeuvreFiche.Panels(4).Text = "CREATION") Or (stb_OeuvreFiche.Panels(4).Text = "MODIFICATION") Then
        stb_OeuvreFiche.SimpleText = "Entrez le code postal de l'association"
    End If
End Sub

Private Sub Txt_Nom_Gotfocus()
    If (stb_OeuvreFiche.Panels(4).Text = "CREATION") Or (stb_OeuvreFiche.Panels(4).Text = "MODIFICATION") Then
        stb_OeuvreFiche.SimpleText = "Entrez le nom de l'association"
    End If
End Sub

Private Sub Txt_nomres_GotFocus()
    If (stb_OeuvreFiche.Panels(4).Text = "CREATION") Or (stb_OeuvreFiche.Panels(4).Text = "MODIFICATION") Then
        stb_OeuvreFiche.SimpleText = "Entrez le nom du responsable de l'association"
    End If
End Sub

Private Sub Txt_rem_GotFocus()
    If (stb_OeuvreFiche.Panels(4).Text = "CREATION") Or (stb_OeuvreFiche.Panels(4).Text = "MODIFICATION") Then
        stb_OeuvreFiche.SimpleText = "Entrez les remarques concernant l'association"
    End If
End Sub

Private Sub Txt_Tel_Gotfocus()
    If (stb_OeuvreFiche.Panels(4).Text = "CREATION") Or (stb_OeuvreFiche.Panels(4).Text = "MODIFICATION") Then
        stb_OeuvreFiche.SimpleText = "Entrez le numéro de téléphone de l'association"
    End If
End Sub

Private Sub txt_Ville_Gotfocus()
    If (stb_OeuvreFiche.Panels(4).Text = "CREATION") Or (stb_OeuvreFiche.Panels(4).Text = "MODIFICATION") Then
        stb_OeuvreFiche.SimpleText = "Entrez la ville de l'association"
    End If
End Sub


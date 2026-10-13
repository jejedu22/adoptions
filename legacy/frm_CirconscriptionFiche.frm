VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_CirconscriptionFiche 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Circonscription"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   ControlBox      =   0   'False
   Icon            =   "frm_CirconscriptionFiche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   5760
         Picture         =   "frm_CirconscriptionFiche.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   6480
         Picture         =   "frm_CirconscriptionFiche.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Frame Fra_circonscriptions 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7335
      Begin VB.TextBox Txt_responsable 
         Height          =   285
         Left            =   2040
         MaxLength       =   40
         TabIndex        =   15
         ToolTipText     =   "Libellé de la circonscription"
         Top             =   3360
         Width           =   3615
      End
      Begin VB.TextBox Txt_ville 
         Height          =   285
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   13
         ToolTipText     =   "Libellé de la circonscription"
         Top             =   2760
         Width           =   3615
      End
      Begin VB.TextBox Txt_cp 
         Height          =   285
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   11
         ToolTipText     =   "Libellé de la circonscription"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Txt_adresse 
         Height          =   285
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   9
         ToolTipText     =   "Libellé de la circonscription"
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox Txt_nom 
         Height          =   285
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   3
         ToolTipText     =   "Libellé de la circonscription"
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox Txt_code 
         Height          =   285
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   2
         ToolTipText     =   "Code de la circonscription"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Responsable"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Ville"
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Code postal"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Adresse"
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lbl_nom 
         Caption         =   "Nom"
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lbl_code 
         Caption         =   "Code"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar stb_CirconscriptionsFiche 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4665
      Width           =   7560
      _ExtentX        =   13335
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
            Object.Width           =   7109
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_CirconscriptionFiche"
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
  
  Me.MousePointer = vbHourglass
  Select Case stb_CirconscriptionsFiche.Panels(4).Text
      Case "CREATION"
          requete = "INSERT INTO CIRC VALUES("
          requete = requete + Iifent(Txt_Code) + ","
          requete = requete + Iifch(Txt_Nom) + ","
          requete = requete + Iifch(Txt_adresse) + ","
          requete = requete + Iifch(Txt_CP) + ","
          requete = requete + Iifch(Txt_Ville) + ","
          requete = requete + Iifch(Txt_responsable) + ")"
          
      Case "MODIFICATION"
          requete = "UPDATE CIRC"
          requete = requete + " SET CIRC_NOM=" + Iifch(Txt_Nom)
          requete = requete + ",CIRC_ADR=" + Iifch(Txt_adresse)
          requete = requete + ",CIRC_CP=" + Iifch(Txt_CP)
          requete = requete + ",CIRC_VILLE=" + Iifch(Txt_Ville)
          requete = requete + ",CIRC_RESP=" + Iifch(Txt_responsable)
          requete = requete + " WHERE CIRC_NUM=" + Iifent(Txt_Code)
          
      Case "SUPPRESSION"
          requete = "DELETE * FROM CIRC WHERE CIRC_NUM = " + Iifent(Txt_Code)
  End Select
  
  Workspaces(0).Databases(0).Execute requete, dbFailOnError
  Me.MousePointer = vbDefault
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
    Me.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub Form_Load()
    QuitterCroix = True
    If frm_Circonscription.stb_ListeCirconscriptions.Panels(4) = "CREATION" Then
        stb_CirconscriptionsFiche.Panels(4).Text = "CREATION"
    End If
    If frm_Circonscription.stb_ListeCirconscriptions.Panels(4) = "MODIFICATION" Then
        stb_CirconscriptionsFiche.Panels(4).Text = "MODIFICATION"
    End If
    If frm_Circonscription.stb_ListeCirconscriptions.Panels(4) = "SUPPRESSION" Then
        stb_CirconscriptionsFiche.Panels(4).Text = "SUPPRESSION"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If QuitterCroix = True Then
    '    If MsgBox("Voulez-vous enregistrer les modifications?", vbYesNo, "ATTENTION !") = vbYes Then
    '        cmd_Valider_Click
    '    Else
    '        cmd_Quitter_Click
    '    End If
    'End If
End Sub

Private Sub Txt_code_Gotfocus()
    If (stb_CirconscriptionsFiche.Panels(4).Text = "CREATION") Or (stb_CirconscriptionsFiche.Panels(4).Text = "MODIFICATION") Then
        stb_CirconscriptionsFiche.SimpleText = "Entrez le code de la circonscription"
    End If
End Sub

Private Sub Txt_Nom_Gotfocus()
    If (stb_CirconscriptionsFiche.Panels(4).Text = "CREATION") Or (stb_CirconscriptionsFiche.Panels(4).Text = "MODIFICATION") Then
        stb_CirconscriptionsFiche.SimpleText = "Entrez le nom de la circonscription"
    End If
End Sub

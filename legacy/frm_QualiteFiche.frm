VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_QualiteFiche 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Qualité"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   ControlBox      =   0   'False
   Icon            =   "frm_QualiteFiche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   4320
         Picture         =   "frm_QualiteFiche.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   4920
         Picture         =   "frm_QualiteFiche.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Frame Fra_serv 
      Caption         =   "Qualités des intervenants sociaux"
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   5535
      Begin VB.TextBox Txt_nom 
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   1
         ToolTipText     =   "Libellé de la qualité"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox Txt_code 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         ToolTipText     =   "Code de la qualité"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lbl_nom 
         Caption         =   "Libellé"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lbl_code 
         Caption         =   "Code"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar stb_QualiteFiche 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2745
      Width           =   5805
      _ExtentX        =   10239
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
            Object.Width           =   4013
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_QualiteFiche"
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
  Select Case stb_QualiteFiche.Panels(4).Text
      Case "CREATION"
          requete = "INSERT INTO QUAL(QUAL_CODE,QUAL_LIB)"
          requete = requete + "VALUES("
          requete = requete + Iifent(Txt_Code) + ","
          requete = requete + Iifch(Txt_Nom) + ");"
      Case "MODIFICATION"
          requete = "UPDATE QUAL"
          requete = requete + " SET QUAL_LIB=" + Iifch(Txt_Nom)
          requete = requete + " WHERE QUAL_CODE=" + Iifent(Txt_Code)
      Case "SUPPRESSION"
          requete = "DELETE * FROM QUAL WHERE QUAL_CODE = " + Iifent(Txt_Code)
  End Select
  Workspaces(0).Databases(0).Execute requete, dbFailOnError
  Unload Me
  frm_Qualite.Dat_Qualite.Refresh
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
    If frm_Qualite.stb_Qualité.Panels(4) = "CREATION" Then
        stb_QualiteFiche.Panels(4).Text = "CREATION"
    End If
    If frm_Qualite.stb_Qualité.Panels(4) = "MODIFICATION" Then
        stb_QualiteFiche.Panels(4).Text = "MODIFICATION"
    End If
    If frm_Qualite.stb_Qualité.Panels(4) = "SUPPRESSION" Then
        stb_QualiteFiche.Panels(4).Text = "SUPPRESSION"
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
    If (stb_QualiteFiche.Panels(4).Text = "CREATION") Or (stb_QualiteFiche.Panels(4).Text = "MODIFICATION") Then
        stb_QualiteFiche.SimpleText = "Entrez le code de la qualité"
    End If
End Sub

Private Sub Txt_Nom_Gotfocus()
    If (stb_QualiteFiche.Panels(4).Text = "CREATION") Or (stb_QualiteFiche.Panels(4).Text = "MODIFICATION") Then
        stb_QualiteFiche.SimpleText = "Entrez le libellé de la qualité"
    End If
End Sub

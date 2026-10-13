VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_CommuneFiche 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commune"
   ClientHeight    =   3450
   ClientLeft      =   540
   ClientTop       =   1515
   ClientWidth     =   7680
   ControlBox      =   0   'False
   Icon            =   "frm_CommuneFiche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3450
   ScaleWidth      =   7680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   7455
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   6720
         Picture         =   "frm_CommuneFiche.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   6000
         Picture         =   "frm_CommuneFiche.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_CommuneFiche 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3195
      Width           =   7680
      _ExtentX        =   13547
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
            TextSave        =   "17/11/99"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11:30"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7320
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Fra_communes 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   7455
      Begin VB.TextBox Txt_num_insee 
         Height          =   285
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   10
         ToolTipText     =   "Numéro INSEE de la commune"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Txt_nom_com 
         Height          =   285
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   9
         ToolTipText     =   "Nom de la commune"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox Txt_cp_com 
         Height          =   285
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   8
         ToolTipText     =   "Code postal de la commune"
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label lbl_cp 
         Caption         =   "Code Postal"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lbl_libelle 
         Caption         =   "Numéro INSEE"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lbl_nom 
         Caption         =   "Nom de la commune"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_CommuneFiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private QuitterCroix As Boolean

Private Sub cmd_Quitter_Click()
    QuitterCroix = False
    Unload Me
End Sub

Private Sub Form_Load()
    QuitterCroix = True
    If frm_Commune.stb_ListeCommune.Panels(4) = "CREATION" Then
        stb_CommuneFiche.Panels(4).Text = "CREATION"
    End If
    If frm_Commune.stb_ListeCommune.Panels(4) = "MODIFICATION" Then
        stb_CommuneFiche.Panels(4).Text = "MODIFICATION"
    End If
    If frm_Commune.stb_ListeCommune.Panels(4) = "SUPPRESSION" Then
        stb_CommuneFiche.Panels(4).Text = "SUPPRESSION"
    End If
End Sub

Private Sub Txt_bureau_GotFocus()
    If (stb_CommuneFiche.Panels(4).Text = "CREATION") Or (stb_CommuneFiche.Panels(4).Text = "MODIFICATION") Then
        stb_CommuneFiche.SimpleText = "Entrez le nom du bureau distributeur"
    End If
End Sub

Private Sub txt_code_depart_GotFocus()
    If (stb_CommuneFiche.Panels(4).Text = "CREATION") Or (stb_CommuneFiche.Panels(4).Text = "MODIFICATION") Then
        stb_CommuneFiche.SimpleText = "Entrez le numéro de département"
    End If
End Sub

Private Sub Txt_cp_com_GotFocus()
    If (stb_CommuneFiche.Panels(4).Text = "CREATION") Or (stb_CommuneFiche.Panels(4).Text = "MODIFICATION") Then
        stb_CommuneFiche.SimpleText = "Entrez le code postal de la commune"
    End If
End Sub

Private Sub txt_maire_GotFocus()
    If (stb_CommuneFiche.Panels(4).Text = "CREATION") Or (stb_CommuneFiche.Panels(4).Text = "MODIFICATION") Then
        stb_CommuneFiche.SimpleText = "Entrez le nom du maire de la commune"
    End If
End Sub

Private Sub Txt_nom_com_GotFocus()
    If (stb_CommuneFiche.Panels(4).Text = "CREATION") Or (stb_CommuneFiche.Panels(4).Text = "MODIFICATION") Then
        stb_CommuneFiche.SimpleText = "Entrez le nom de la commune"
    End If
End Sub

Private Sub Txt_num_insee_GotFocus()
    If (stb_CommuneFiche.Panels(4).Text = "CREATION") Or (stb_CommuneFiche.Panels(4).Text = "MODIFICATION") Then
        stb_CommuneFiche.SimpleText = "Entrez le numéro INSEE de la commune"
    End If
End Sub

Private Sub txt_telephone_GotFocus()
    If (stb_CommuneFiche.Panels(4).Text = "CREATION") Or (stb_CommuneFiche.Panels(4).Text = "MODIFICATION") Then
        stb_CommuneFiche.SimpleText = "Entrez le numéro téléphone de la mairie"
    End If
End Sub

Private Sub cmd_Valider_Click()

  Dim requete As String
      
  QuitterCroix = False
  On Error GoTo Rupture_Integrite
  
  Select Case stb_CommuneFiche.Panels(4).Text
    Case "CREATION"
        requete = "INSERT INTO COMM VALUES ("
        requete = requete + Iifent(Txt_num_insee) + ","
        requete = requete + Iifch(Txt_nom_com) + ","
        requete = requete + Iifch(Txt_cp_com) + ")"
        
    Case "MODIFICATION"
        requete = "UPDATE COMM"
        requete = requete + " SET COMM_NOM=" + Iifch(Txt_nom_com)
        requete = requete + ",COMM_CP=" + Iifch(Txt_cp_com)
        requete = requete + " WHERE COMM_INSEE=" + Iifent(Txt_num_insee)
        
    Case "SUPPRESSION"
        requete = "DELETE * FROM COMM WHERE COMM_INSEE = " + Iifent(Txt_num_insee)
  End Select
  
  Workspaces(0).Databases(0).Execute requete, dbFailOnError
  Unload Me
  Exit Sub
    
Rupture_Integrite:
    
  Select Case Err
  Case 3022
    MsgBox "Cette commune existe déjà, création impossible", vbOKOnly + vbExclamation
  Case 3200
    MsgBox "Cette commune est référencée dans une autre table, suppression impossible", vbOKOnly + vbExclamation
  Case 3075
    MsgBox "l'un des champs contient une cote, mise à jour impossible", vbOKOnly + vbExclamation
  Case 3315
    MsgBox "Un champs obligatoire n'est pas renseigné, mise à jour impossible", vbOKOnly + vbExclamation
  Case Else
    MsgBox "Erreur " + CStr(Err) + Err.Description + " survenue, mise à jour impossible", vbOKOnly + vbExclamation
  End Select
  
  Exit Sub

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


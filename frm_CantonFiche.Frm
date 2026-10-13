VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_CantonFiche 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Canton"
   ClientHeight    =   3660
   ClientLeft      =   540
   ClientTop       =   1515
   ClientWidth     =   6840
   ControlBox      =   0   'False
   Icon            =   "frm_CantonFiche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3660
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   5280
         Picture         =   "frm_CantonFiche.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   6000
         Picture         =   "frm_CantonFiche.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Frame Fra_cantons 
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   6615
      Begin VB.TextBox Txt_code_canton 
         Height          =   285
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   8
         ToolTipText     =   "Numéro du canton"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Txt_nom_canton 
         Height          =   285
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   7
         ToolTipText     =   "Nom du canton"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label lbl_libelle 
         Caption         =   "Nom du canton"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lbl_code 
         Caption         =   "Code du canton"
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
   End
   Begin MSComctlLib.StatusBar stb_CantonFiche 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3405
      Width           =   6840
      _ExtentX        =   12065
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
            Object.Width           =   5839
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_CantonFiche"
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
    If frm_Canton.stb_ListeCanton.Panels(4) = "CREATION" Then
        stb_CantonFiche.Panels(4).Text = "CREATION"
    End If
    If frm_Canton.stb_ListeCanton.Panels(4) = "MODIFICATION" Then
        stb_CantonFiche.Panels(4).Text = "MODIFICATION"
    End If
    If frm_Canton.stb_ListeCanton.Panels(4) = "SUPPRESSION" Then
        stb_CantonFiche.Panels(4).Text = "SUPPRESSION"
    End If
End Sub

Private Sub Txt_Code_canton_GotFocus()
    If (stb_CantonFiche.Panels(4).Text = "CREATION") Or (stb_CantonFiche.Panels(4).Text = "MODIFICATION") Then
        stb_CantonFiche.SimpleText = "Entrez le code du canton"
    End If
End Sub

Private Sub txt_code_depart_GotFocus()
    If (stb_CantonFiche.Panels(4).Text = "CREATION") Or (stb_CantonFiche.Panels(4).Text = "MODIFICATION") Then
        stb_CantonFiche.SimpleText = "Entrez le code du département"
    End If
End Sub

Private Sub Txt_Nom_canton_GotFocus()
    If (stb_CantonFiche.Panels(4).Text = "CREATION") Or (stb_CantonFiche.Panels(4).Text = "MODIFICATION") Then
        stb_CantonFiche.SimpleText = "Entrez le nom du canton"
    End If
End Sub

Private Sub cmd_Valider_Click()
    Dim requete As String
        
    QuitterCroix = False
    On Error GoTo Rupture_Integrite
    
    Select Case stb_CantonFiche.Panels(4).Text
      Case "CREATION"
        requete = "INSERT INTO CANT VALUES ("
        requete = requete + Iifent(Txt_code_canton) + ","
        requete = requete + Iifch(Txt_nom_canton) + ")"
        
      Case "MODIFICATION"
        requete = "UPDATE CANT"
        requete = requete + " SET CANT_NOM=" + Iifch(Txt_nom_canton)
        requete = requete + " WHERE CANT_NUM=" + Iifent(Txt_code_canton)
        
      Case "SUPPRESSION"
        requete = "DELETE * FROM CANT WHERE CANT_NUM = " + Iifent(Txt_code_canton)
    End Select
    
    Workspaces(0).Databases(0).Execute requete, dbFailOnError
    Unload Me
    frm_Canton.Dat_Canton.Refresh
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If QuitterCroix = True Then
    '    If MsgBox("Voulez-vous enregistrer les modifications?", vbYesNo, "ATTENTION !") = vbYes Then
    '        cmd_Valider_Click
    '    Else
    '        cmd_Quitter_Click
    '    End If
    'End If
End Sub


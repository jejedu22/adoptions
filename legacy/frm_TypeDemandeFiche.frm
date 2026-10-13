VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_TypeDemandeFiche 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Type des demandes"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   ControlBox      =   0   'False
   Icon            =   "frm_TypeDemandeFiche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   5280
         Picture         =   "frm_TypeDemandeFiche.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   5880
         Picture         =   "frm_TypeDemandeFiche.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Frame Fra_Nature 
      Caption         =   "Type de demande de placements"
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   6495
      Begin VB.TextBox Txt_nom 
         Height          =   285
         Left            =   1560
         MaxLength       =   31
         TabIndex        =   1
         ToolTipText     =   "Libellé du type de la demande"
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox Txt_code 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         ToolTipText     =   "Code du type de demande"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lbl_libelletype 
         Caption         =   "Libellé"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lbl_code 
         Caption         =   "Code"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar stb_TypeFiche 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2775
      Width           =   6735
      _ExtentX        =   11880
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
            Object.Width           =   5653
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_TypeDemandeFiche"
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
    Select Case stb_TypeFiche.Panels(4).Text
        Case "CREATION"
            requete = "INSERT INTO TYPE(TYPE_CODE,TYPE_LIBELLE)"
            requete = requete + "VALUES("
            requete = requete + Iifent(Txt_Code) + ","
            requete = requete + Iifch(Txt_Nom) + ");"
        Case "MODIFICATION"
            requete = "UPDATE TYPE"
            requete = requete + " SET TYPE_LIBELLE=" + Iifch(Txt_Nom)
            requete = requete + " WHERE TYPE_CODE=" + Iifent(Txt_Code)
        Case "SUPPRESSION"
            requete = "DELETE * FROM TYPE WHERE TYPE_CODE = " + Iifent(Txt_Code)
    End Select
    Workspaces(0).Databases(0).Execute requete, dbFailOnError
    Unload Me
    frm_TypeDemande.Dat_Types.Refresh
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
    If frm_TypeDemande.stb_Types.Panels(4) = "CREATION" Then
        stb_TypeFiche.Panels(4).Text = "CREATION"
    End If
    If frm_TypeDemande.stb_Types.Panels(4) = "MODIFICATION" Then
        stb_TypeFiche.Panels(4).Text = "MODIFICATION"
    End If
    If frm_TypeDemande.stb_Types.Panels(4) = "SUPPRESSION" Then
        stb_TypeFiche.Panels(4).Text = "SUPPRESSION"
    End If
End Sub

Private Sub Txt_code_Gotfocus()
    If (stb_TypeFiche.Panels(4).Text = "CREATION") Or (stb_TypeFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TypeFiche.SimpleText = "Entrez le code du type de demande"
    End If
End Sub

Private Sub Txt_Nom_Gotfocus()
    If (stb_TypeFiche.Panels(4).Text = "CREATION") Or (stb_TypeFiche.Panels(4).Text = "MODIFICATION") Then
        stb_TypeFiche.SimpleText = "Entrez le libellé du type de la demande"
    End If
End Sub




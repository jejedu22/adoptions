VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FamilleDoc_Liste 
   Caption         =   "Liste des familles de documents"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   ControlBox      =   0   'False
   Icon            =   "FamilleDoc_Liste.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   70
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8040
         Picture         =   "FamilleDoc_Liste.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Valider"
         Top             =   150
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Data Dat_fam 
         Caption         =   "Dat_fam"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "FamilleDoc_Liste.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Modifier la famille de documents sélectionnée"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "FamilleDoc_Liste.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Créer une famille de documents"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "FamilleDoc_Liste.frx":0748
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Supprimer la famille de documents sélectionnée"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8760
         Picture         =   "FamilleDoc_Liste.frx":084A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_ListeFamilles 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4830
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   450
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
            TextSave        =   "11:27"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8969
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin MSDBGrid.DBGrid DBG_famille 
      Bindings        =   "FamilleDoc_Liste.frx":094C
      Height          =   3615
      Left            =   1320
      OleObjectBlob   =   "FamilleDoc_Liste.frx":0962
      TabIndex        =   6
      ToolTipText     =   "Liste des sources de données"
      Top             =   960
      Width           =   6375
   End
End
Attribute VB_Name = "FamilleDoc_Liste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Creer_Click()

  stb_ListeFamilles.Panels(4).Text = "CREATION"
  Load FamilleDoc_Fiche
  FamilleDoc_Fiche.Caption = "Création d'une famille de documents"
  If Dat_fam.Recordset.EOF Then
    FamilleDoc_Fiche.Txt_id = 1
  Else
    With Dat_fam.Recordset
      .MoveFirst
      .MoveLast
    End With
    FamilleDoc_Fiche.Txt_id = Dat_fam.Recordset.RecordCount + 1
  End If
  FamilleDoc_Fiche.Show vbModal
  stb_ListeFamilles.Panels(4).Text = ""

End Sub

Private Sub cmd_Modifier_Click()

  If Not Dat_fam.Recordset.EOF Then
    If Dat_fam.Recordset("wfam_bloque") = 1 Then
      MsgBox "Vous ne pouvez pas modifer cette famille de documents.", vbExclamation, "IMPOSSIBLE"
    Else
      stb_ListeFamilles.Panels(4).Text = "MODIFICATION"
      Load FamilleDoc_Fiche
      With FamilleDoc_Fiche
        .Caption = "Modification d'une famille de documents"
        .Txt_id = Dat_fam.Recordset("wfam_id")
        .Txt_Nom = Dat_fam.Recordset("wfam_nom")
        .Cmb_Bloque.ListIndex = 1 'non
        .Show vbModal
        stb_ListeFamilles.Panels(3).Text = ""
      End With
    End If
  End If
  
End Sub


Private Sub cmd_Quitter_Click()
  Unload Me
End Sub

Private Sub cmd_Supprimer_Click()

  If Not Dat_fam.Recordset.EOF Then
    If Dat_fam.Recordset("wfam_bloque") = 1 Then
        MsgBox "Vous ne pouvez pas supprimer cette famille de documents.", vbExclamation, "IMPOSSIBLE"
    Else
      stb_ListeFamilles.Panels(4).Text = "SUPPRESSION"
      Load FamilleDoc_Fiche
      With FamilleDoc_Fiche
        .Caption = "Suppression d'une famille de documents"
        .Txt_id = Dat_fam.Recordset("wfam_id")
        .Txt_Nom = Dat_fam.Recordset("wfam_nom")
        .Txt_Nom.Enabled = False
        .Cmb_Bloque.ListIndex = 1 'non
        .Show vbModal
        stb_ListeFamilles.Panels(3).Text = ""
      End With
    End If
  End If
  
End Sub

Private Sub cmd_Valider_Click()

  If Not Dat_fam.Recordset.EOF Then
    DocumentWord_fiche.txt_IdFam = Dat_fam.Recordset("wfam_id")
    DocumentWord_fiche.Txt_Famille = Dat_fam.Recordset("wfam_nom")
  End If
  Unload Me

End Sub

Private Sub DBG_famille_HeadClick(ByVal ColIndex As Integer)

  Dim colonne As String
  Dim classem As String
  
  colonne = DBG_famille.Columns(ColIndex).DataField
  Call Tri(Dat_fam, DBG_famille, colonne)
      
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_fam.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_ListeFamilles.Panels(4).Text = "Classement " & classem

End Sub


Private Sub Form_Load()

  Dat_fam.DatabaseName = CheminBase
  With Dat_fam
    .RecordSource = "select * from wfam"
    .Refresh
  End With
    
End Sub



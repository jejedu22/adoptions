VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SourcesDonnées_liste 
   Caption         =   "Liste des sources de données"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   ControlBox      =   0   'False
   Icon            =   "SourcesDonnées_Liste.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Dat_source 
      Caption         =   "Dat_source"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3765
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8040
         Picture         =   "SourcesDonnées_Liste.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Valider"
         Top             =   150
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8760
         Picture         =   "SourcesDonnées_Liste.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "SourcesDonnées_Liste.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprimer la source de données"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "SourcesDonnées_Liste.frx":0748
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Créer une nouvelle source de données"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "SourcesDonnées_Liste.frx":084A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Modifier la source de données"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_ListeSources 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   5610
      Width           =   9495
      _ExtentX        =   16748
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
            TextSave        =   "11:57"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9102
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin MSDBGrid.DBGrid DBG_source 
      Bindings        =   "SourcesDonnées_Liste.frx":094C
      Height          =   4215
      Left            =   600
      OleObjectBlob   =   "SourcesDonnées_Liste.frx":0965
      TabIndex        =   6
      ToolTipText     =   "Liste des sources de données"
      Top             =   960
      Width           =   8295
   End
End
Attribute VB_Name = "SourcesDonnées_liste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Creer_Click()

  stb_ListeSources.Panels(3).Text = "CREATION"
  Load SourcesDonnées_fiche
  SourcesDonnées_fiche.Caption = "Création d'une source de données"

  If Dat_source.Recordset.EOF Then
    SourcesDonnées_fiche.Txt_id = 1
  Else
    SourcesDonnées_fiche.Dat_Id.RecordSource = "select max(wsou_id) from wsou"
    SourcesDonnées_fiche.Dat_Id.Refresh
    SourcesDonnées_fiche.Txt_id.Text = SourcesDonnées_fiche.Dat_Id.Recordset(0) + 1
    SourcesDonnées_fiche.Txt_nom = ""
  End If
  
  With SourcesDonnées_fiche
    .Cmd_colonnes.Enabled = False
    .Cmd_Conditions.Enabled = False
    .Cmd_ordre.Enabled = False
  End With
  
  SourcesDonnées_fiche.Show vbModal
  stb_ListeSources.Panels(3).Text = ""
  
End Sub

Private Sub cmd_Modifier_Click()

  If Not Dat_source.Recordset.EOF Then
    If Dat_source.Recordset("wsou_bloque") = 1 Then
      MsgBox "Vous ne pouvez pas modifier cette source de données.", vbExclamation, "IMPOSSIBLE"
    Else
      stb_ListeSources.Panels(3).Text = "MODIFICATION"
      Load SourcesDonnées_fiche
      SourcesDonnées_fiche.Caption = "Modification d'une source de données"
      SourcesDonnées_fiche.Txt_id = Dat_source.Recordset("wsou_id")
      SourcesDonnées_fiche.Txt_nom = Dat_source.Recordset("wsou_nom")
      SourcesDonnées_fiche.Cmb_Vue.Text = Dat_source.Recordset("wsou_vue")
      SourcesDonnées_fiche.Cmb_Vue.Enabled = False
      SourcesDonnées_fiche.Cmb_Bloque.ListIndex = 1 'non
      
      SourcesDonnées_fiche.Show vbModal
      stb_ListeSources.Panels(3).Text = ""
    End If
  End If
      
End Sub


Private Sub cmd_Quitter_Click()
  
  Unload Me
    
End Sub

Private Sub cmd_Supprimer_Click()
  
Dim sql As String
    
  If Not Dat_source.Recordset.EOF Then
    If Dat_source.Recordset("wsou_bloque") = 1 Then
      MsgBox "Vous ne pouvez pas supprimer cette source de données.", vbExclamation, "IMPOSSIBLE"
    Else
      stb_ListeSources.Panels(3).Text = "SUPPRESSION"
      Load SourcesDonnées_fiche
      SourcesDonnées_fiche.Caption = "Suppression d'une source de données"
      SourcesDonnées_fiche.Txt_id.Enabled = False
      SourcesDonnées_fiche.Txt_nom.Enabled = False
      SourcesDonnées_fiche.Cmb_Bloque.Enabled = False
      SourcesDonnées_fiche.Cmb_Bloque.ListIndex = 1
      SourcesDonnées_fiche.Cmb_Vue.Text = Dat_source.Recordset("wsou_vue")
      SourcesDonnées_fiche.Cmb_Vue.Enabled = False
      SourcesDonnées_fiche.Cmd_colonnes.Enabled = False
      SourcesDonnées_fiche.Cmd_ordre.Enabled = False
      SourcesDonnées_fiche.Cmd_Conditions.Enabled = False
      
      SourcesDonnées_fiche.Txt_id = Dat_source.Recordset("wsou_id")
      SourcesDonnées_fiche.Txt_nom = Dat_source.Recordset("wsou_nom")
      SourcesDonnées_fiche.Show vbModal
      stb_ListeSources.Panels(3).Text = ""
    End If
  End If
  
End Sub


Private Sub cmd_Valider_Click()

  If Not Dat_source.Recordset.EOF Then
    DocumentWord_fiche.Txt_IdSource = Dat_source.Recordset("wsou_id")
    DocumentWord_fiche.Txt_Source = Dat_source.Recordset("wsou_nom")
  End If
  Unload Me
  
End Sub

Private Sub DBG_source_HeadClick(ByVal ColIndex As Integer)

  Dim colonne As String
  Dim classem As String
  
  colonne = DBG_source.Columns(ColIndex).DataField
  Call Tri(Dat_source, DBG_source, colonne)
      
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_source.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_ListeSources.Panels(4).Text = "Classement " & classem

End Sub


Private Sub Form_Load()

  Dat_source.DatabaseName = CheminBase
  Dat_source.RecordSource = "select * from WSOU order by wsou_id"
  Dat_source.Refresh
  
End Sub



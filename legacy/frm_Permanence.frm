VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Permanence 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Permanences des intervenants sociaux"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_Permanence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Dat_Perm 
      Caption         =   "Dat_Perm"
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
      RecordSource    =   "PERM"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Actualiser 
         Default         =   -1  'True
         Height          =   375
         Left            =   7320
         Picture         =   "frm_Permanence.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Rechercher d'après les critères"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8040
         Picture         =   "frm_Permanence.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8760
         Picture         =   "frm_Permanence.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Excel 
         Height          =   375
         Left            =   4560
         Picture         =   "frm_Permanence.frx":0748
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Export EXCEL"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "frm_Permanence.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprimer la permanence sélectionnée"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "frm_Permanence.frx":09D4
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Créer une nouvelle permanence"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "frm_Permanence.frx":0AD6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modifier les données de la permanence sélectionnée"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   9255
      Begin VB.TextBox txt_Ville 
         Height          =   285
         Left            =   5520
         MaxLength       =   30
         TabIndex        =   9
         Top             =   600
         Width           =   3495
      End
      Begin VB.CommandButton cmd_Ville 
         Caption         =   "&Ville"
         Height          =   255
         Left            =   3960
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Txt_lieu 
         Height          =   285
         Left            =   5520
         MaxLength       =   30
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox txt_code 
         Height          =   285
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   7
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton cmd_lieu 
         Caption         =   "&Lieu"
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_code 
         Caption         =   "&Code"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.StatusBar stb_Perm 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   6570
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "30/11/99"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "09:31"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9578
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin MSDBGrid.DBGrid grd_Perm 
      Bindings        =   "frm_Permanence.frx":0BD8
      Height          =   4815
      Left            =   120
      OleObjectBlob   =   "frm_Permanence.frx":0BEF
      TabIndex        =   10
      Top             =   1680
      Width           =   9255
   End
End
Attribute VB_Name = "frm_Permanence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Actualiser_Click()
    Dim sel_code As String
    Dim sel_lieu As String
    Dim sel_ville As String
    
    If Txt_Code <> "" Then sel_code = " AND PERM_CODE like """ & Txt_Code & "*"""
    If Txt_Lieu <> "" Then sel_lieu = " AND PERM_LIEU like """ & Txt_lib & "*"""
    If Txt_Ville <> "" Then sel_ville = " AND PERM_VILLE like """ & Txt_Ville & "*"""
        
    Dat_Perm.RecordSource = "SELECT * FROM PERM WHERE TRUE" + _
        sel_code + _
        sel_lieu + _
        sel_ville + _
        ClauseOrder
        
    Me.MousePointer = vbHourglass
    Dat_Perm.Refresh
    Me.MousePointer = vbDefault
End Sub

Private Sub cmd_code_Click()
    Txt_Code = ""
    Txt_Code.SetFocus
End Sub

Private Sub cmd_Creer_Click()
    stb_Perm.Panels(4).Text = "CREATION"
    Load frm_PermanenceFiche
    frm_PermanenceFiche.Caption = "Création d'une permanence"
    frm_PermanenceFiche.Txt_Code.Enabled = False
    frm_PermanenceFiche.Txt_Code = NouveauCode("PERM", "PERM_CODE")
    frm_PermanenceFiche.Txt_Lieu = ""
    frm_PermanenceFiche.Txt_Ville = ""
    frm_PermanenceFiche.Show vbModal
  
    stb_Perm.Panels(4).Text = ""
End Sub

Private Sub cmd_Excel_Click()

  ExportExcel (Dat_Perm.RecordSource)
  
End Sub

Private Sub cmd_lieu_Click()
    Txt_Lieu = ""
    Txt_Lieu.SetFocus
End Sub

Private Sub cmd_Modifier_Click()
    If Not Dat_Perm.Recordset.EOF Then
        stb_Perm.Panels(4).Text = "MODIFICATION"
        Load frm_PermanenceFiche
        frm_PermanenceFiche.Caption = "Modification d'une permanence"
        frm_PermanenceFiche.Txt_Code.Enabled = False
        frm_PermanenceFiche.Txt_Code = Dat_Perm.Recordset("PERM_CODE")
        If Dat_Perm.Recordset("PERM_LIEU") <> "" Then
            frm_PermanenceFiche.Txt_Lieu = Dat_Perm.Recordset("PERM_LIEU")
        Else
            frm_PermanenceFiche.Txt_Lieu = ""
        End If
        If Dat_Perm.Recordset("PERM_VILLE") <> "" Then
            frm_PermanenceFiche.Txt_Ville = Dat_Perm.Recordset("PERM_VILLE")
        Else
            frm_PermanenceFiche.Txt_Ville = ""
        End If
        frm_PermanenceFiche.Show vbModal
        
        stb_Perm.Panels(4).Text = ""
    End If
End Sub

Private Sub cmd_Quitter_Click()
    Unload Me
End Sub

Private Sub cmd_Supprimer_Click()
    If Not Dat_Perm.Recordset.EOF Then
        stb_Perm.Panels(4).Text = "SUPPRESSION"
        Load frm_PermanenceFiche
        frm_PermanenceFiche.Caption = "Suppression d'une permanence"
        frm_PermanenceFiche.Txt_Code.Enabled = False
        frm_PermanenceFiche.Txt_Lieu.Enabled = False
        frm_PermanenceFiche.Txt_Ville.Enabled = False

        frm_PermanenceFiche.Txt_Code = Dat_Perm.Recordset("PERM_CODE")
        If Dat_Perm.Recordset("PERM_LIEU") <> "" Then
            frm_PermanenceFiche.Txt_Lieu = Dat_Perm.Recordset("PERM_LIEU")
        Else
            frm_PermanenceFiche.Txt_Lieu = ""
        End If
        If Dat_Perm.Recordset("PERM_VILLE") <> "" Then
            frm_PermanenceFiche.Txt_Ville = Dat_Perm.Recordset("PERM_VILLE")
        Else
            frm_PermanenceFiche.Txt_Ville = ""
        End If
        frm_PermanenceFiche.Show vbModal
        
        stb_Perm.Panels(4).Text = ""
    End If
End Sub

Private Sub cmd_Valider_Click()
    If Not Dat_Perm.Recordset.EOF Then
        Me.Hide
    End If
End Sub

Private Sub cmd_Ville_Click()
    Txt_Ville = ""
    Txt_Ville.SetFocus
End Sub

Private Sub Form_Load()
    Dat_Perm.DatabaseName = CheminBase
    Cmd_Actualiser_Click
End Sub

Private Sub grd_Perm_HeadClick(ByVal ColIndex As Integer)

  Dim colonne As String
  Dim classem As String
  
  colonne = grd_Perm.Columns(ColIndex).DataField
  Call Tri(Dat_Perm, grd_Perm, colonne)
    
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_Perm.RecordSource, "ASC") <> 0 Then
    classem = "Croissant"
  Else
    classem = "Décroissant"
  End If
       
  stb_Perm.Panels(4).Text = "Classement " & classem
  
End Sub

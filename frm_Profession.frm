VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Profession 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Professions"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   Icon            =   "frm_Profession.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_Outils 
      ClipControls    =   0   'False
      Height          =   600
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Actualiser 
         Default         =   -1  'True
         Height          =   375
         Left            =   7320
         Picture         =   "frm_Profession.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "REchercher d'après les critères"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8040
         Picture         =   "frm_Profession.frx":09CC
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
         Picture         =   "frm_Profession.frx":0ACE
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Excel 
         Height          =   375
         Left            =   4560
         Picture         =   "frm_Profession.frx":0BD0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Export EXCEL"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "frm_Profession.frx":0D5A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprimer la profession sélectionnée"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "frm_Profession.frx":0E5C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Créer une nouvelle profession"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "frm_Profession.frx":0F5E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modifier les données de la profession sélectionnée"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   9255
      Begin VB.TextBox txt_Prof_lib2 
         Height          =   285
         Left            =   5520
         MaxLength       =   30
         TabIndex        =   9
         Top             =   720
         Width           =   3495
      End
      Begin VB.CommandButton cmd_Prof_lib2 
         Caption         =   "Li&bellé 2"
         Height          =   255
         Left            =   3960
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txt_Prof_lib1 
         Height          =   285
         Left            =   5520
         MaxLength       =   30
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox txt_Prof_code 
         Height          =   285
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   7
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton cmd_Prof_lib1 
         Caption         =   "&Libellé 1"
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_Prof_code 
         Caption         =   "&Code"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Data Dat_Prof 
      Caption         =   "Dat_Prof"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PROF"
      Top             =   3315
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSDBGrid.DBGrid grd_Prof 
      Bindings        =   "frm_Profession.frx":1060
      Height          =   4695
      Left            =   120
      OleObjectBlob   =   "frm_Profession.frx":1077
      TabIndex        =   10
      Top             =   1800
      Width           =   9255
   End
   Begin MSComctlLib.StatusBar stb_Profession 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   6540
      Width           =   9480
      _ExtentX        =   16722
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
            TextSave        =   "09:33"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9525
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_Profession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Actualiser_Click()
    Dim sel_code As String
    Dim sel_lib1 As String
    Dim sel_lib2 As String
    
    If txt_Prof_code <> "" Then sel_code = " AND PROF_CODE like """ & txt_Prof_code & "*"""
    If txt_Prof_lib1 <> "" Then sel_lib1 = " AND PROF_DESI1 like """ & txt_Prof_lib1 & "*"""
    If txt_Prof_lib2 <> "" Then sel_lib2 = " AND PROF_DESI2 like """ & txt_Prof_lib2 & "*"""
    
    Dat_Prof.RecordSource = "SELECT * FROM PROF WHERE TRUE" + _
        sel_code + _
        sel_lib1 + _
        sel_lib2 + _
        ClauseOrder
        
    Me.MousePointer = vbHourglass
    Dat_Prof.Refresh
    Me.MousePointer = vbDefault

End Sub

Private Sub cmd_Creer_Click()
    stb_Profession.Panels(4).Text = "CREATION"
    Load frm_ProfessionFiche
    frm_ProfessionFiche.Caption = "Création d'une profession"
    frm_ProfessionFiche.Txt_Code.Enabled = False
    frm_ProfessionFiche.Txt_Code = NouveauCode("PROF", "PROF_CODE")
    frm_ProfessionFiche.Txt_lib1 = ""
    frm_ProfessionFiche.Txt_lib2 = ""
    frm_ProfessionFiche.Show vbModal
  
    stb_Profession.Panels(4).Text = ""
End Sub

Private Sub cmd_Excel_Click()

  ExportExcel (Dat_Prof.RecordSource)

End Sub

Private Sub cmd_Modifier_Click()
    If Not Dat_Prof.Recordset.EOF Then
        stb_Profession.Panels(4).Text = "MODIFICATION"
        Load frm_ProfessionFiche
        frm_ProfessionFiche.Caption = "Modification d'une profession"
        frm_ProfessionFiche.Txt_Code.Enabled = False
        frm_ProfessionFiche.Txt_Code = Dat_Prof.Recordset("PROF_CODE")
        If Dat_Prof.Recordset("PROF_DESI1") <> "" Then
            frm_ProfessionFiche.Txt_lib1 = Dat_Prof.Recordset("PROF_DESI1")
        Else
            frm_ProfessionFiche.Txt_lib1 = ""
        End If
        If Dat_Prof.Recordset("PROF_DESI2") <> "" Then
            frm_ProfessionFiche.Txt_lib2 = Dat_Prof.Recordset("PROF_DESI2")
        Else
            frm_ProfessionFiche.Txt_lib2 = ""
        End If
        frm_ProfessionFiche.Show vbModal
        
        stb_Profession.Panels(4).Text = ""
    End If
End Sub

Private Sub cmd_Prof_code_Click()
    txt_Prof_code = ""
    txt_Prof_code.SetFocus
End Sub

Private Sub cmd_Prof_lib1_Click()
    txt_Prof_lib1 = ""
    txt_Prof_lib1.SetFocus
End Sub

Private Sub cmd_Prof_lib2_Click()
    txt_Prof_lib2 = ""
    txt_Prof_lib2.SetFocus
End Sub

Private Sub cmd_Quitter_Click()
    Unload frm_Profession
End Sub

Private Sub cmd_Supprimer_Click()
    If Not Dat_Prof.Recordset.EOF Then
        stb_Profession.Panels(4).Text = "SUPPRESSION"
        Load frm_ProfessionFiche
        frm_ProfessionFiche.Caption = "Suppression d'une profession"
        frm_ProfessionFiche.Txt_Code.Enabled = False
        frm_ProfessionFiche.Txt_lib1.Enabled = False
        frm_ProfessionFiche.Txt_lib2.Enabled = False
        
        frm_ProfessionFiche.Txt_Code = Dat_Prof.Recordset("PROF_CODE")
        If Dat_Prof.Recordset("PROF_DESI1") <> "" Then
            frm_ProfessionFiche.Txt_lib1 = Dat_Prof.Recordset("PROF_DESI1")
        Else
            frm_ProfessionFiche.Txt_lib1 = ""
        End If
        If Dat_Prof.Recordset("PROF_DESI2") <> "" Then
            frm_ProfessionFiche.Txt_lib2 = Dat_Prof.Recordset("PROF_DESI2")
        Else
            frm_ProfessionFiche.Txt_lib2 = ""
        End If
        frm_ProfessionFiche.Show vbModal
        
        stb_Profession.Panels(4).Text = ""
    End If

End Sub

Private Sub cmd_Valider_Click()
    If Not Dat_Prof.Recordset.EOF Then
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Dat_Prof.DatabaseName = CheminBase
    Cmd_Actualiser_Click
End Sub

Private Sub grd_Prof_HeadClick(ByVal ColIndex As Integer)

  Dim colonne As String
  Dim classem As String
  
  colonne = grd_Prof.Columns(ColIndex).DataField
  Call Tri(Dat_Prof, grd_Prof, colonne)
      
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_Prof.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_Profession.Panels(4).Text = "Classement " & classem

End Sub


VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_OrigineEthnique 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Origines Ethniques"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_OrigineEthnique.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Dat_OriginesEthniques 
      Caption         =   "Dat_OriginesEthniques"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "INST"
      Top             =   6000
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   9255
      Begin VB.CommandButton cmd_Origine_code 
         Caption         =   "&Code"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_Origine_lib 
         Caption         =   "&Libellé"
         Height          =   255
         Left            =   3960
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txt_Origine_code 
         Height          =   285
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   10
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox Txt_Origine_lib 
         Height          =   285
         Left            =   5520
         MaxLength       =   30
         TabIndex        =   11
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "frm_OrigineEthnique.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modifier les données de l'institution sélectionnée"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "frm_OrigineEthnique.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Créer une nouvelle institution"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "frm_OrigineEthnique.frx":0ACE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprimer l'institution sélectionnée"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Excel 
         Height          =   375
         Left            =   4560
         Picture         =   "frm_OrigineEthnique.frx":0BD0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Export EXCEL"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8760
         Picture         =   "frm_OrigineEthnique.frx":0D5A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8040
         Picture         =   "frm_OrigineEthnique.frx":0E5C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Actualiser 
         Default         =   -1  'True
         Height          =   375
         Left            =   7320
         Picture         =   "frm_OrigineEthnique.frx":0F5E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Rechercher d'après les critères"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSDBGrid.DBGrid grd_OriginesEthniques 
      Bindings        =   "frm_OrigineEthnique.frx":1060
      Height          =   5055
      Left            =   2400
      OleObjectBlob   =   "frm_OrigineEthnique.frx":1084
      TabIndex        =   9
      Top             =   1440
      Width           =   4695
   End
   Begin MSComctlLib.StatusBar stb_OriginesEthniques 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
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
            TextSave        =   "09:30"
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
End
Attribute VB_Name = "frm_OrigineEthnique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Actualiser_Click()
    Dim sel_code As String
    Dim sel_lib As String
    
    If txt_Origine_code <> "" Then sel_code = " AND ORET_CODE like """ & txt_Origine_code & "*"""
    If Txt_Origine_lib <> "" Then sel_lib = " AND ORET_LIBELLE like """ & Txt_Origine_lib & "*"""
        
    Dat_OriginesEthniques.RecordSource = "SELECT * FROM ORET WHERE TRUE" + _
        sel_code + _
        sel_lib + _
        ClauseOrder
        
    Me.MousePointer = vbHourglass
    Dat_OriginesEthniques.Refresh
    Me.MousePointer = vbDefault
End Sub

Private Sub cmd_Creer_Click()
    stb_OriginesEthniques.Panels(4).Text = "CREATION"
    Load frm_OrigineEthniqueFiche
    frm_OrigineEthniqueFiche.Caption = "Création d'une origine ethnique"
    frm_OrigineEthniqueFiche.Txt_Code.Enabled = False
    frm_OrigineEthniqueFiche.Txt_Code = NouveauCode("ORET", "ORET_CODE")
    frm_OrigineEthniqueFiche.Txt_Nom = ""
    frm_OrigineEthniqueFiche.Show vbModal
  
    stb_OriginesEthniques.Panels(4).Text = ""
End Sub

Private Sub cmd_Excel_Click()

  ExportExcel (Dat_OriginesEthniques.RecordSource)

End Sub

Private Sub cmd_Modifier_Click()
    If Not Dat_OriginesEthniques.Recordset.EOF Then
        stb_OriginesEthniques.Panels(4).Text = "MODIFICATION"
        Load frm_OrigineEthniqueFiche
        frm_OrigineEthniqueFiche.Caption = "Modification d'une origine ethnique"
        frm_OrigineEthniqueFiche.Txt_Code.Enabled = False
        frm_OrigineEthniqueFiche.Txt_Code = Dat_OriginesEthniques.Recordset("ORET_CODE")
        If Dat_OriginesEthniques.Recordset("ORET_LIBELLE") <> "" Then
            frm_OrigineEthniqueFiche.Txt_Nom = Dat_OriginesEthniques.Recordset("ORET_LIBELLE")
        Else
            frm_OrigineEthniqueFiche.Txt_Nom = ""
        End If
        frm_OrigineEthniqueFiche.Show vbModal
        
        stb_OriginesEthniques.Panels(4).Text = ""
    End If
End Sub

Private Sub cmd_Origine_code_Click()
    txt_Origine_code = ""
    txt_Origine_code.SetFocus
End Sub

Private Sub cmd_Origine_lib_Click()
    Txt_Origine_lib = ""
    Txt_Origine_lib.SetFocus
End Sub

Private Sub cmd_Quitter_Click()
    Unload Me
End Sub

Private Sub cmd_Supprimer_Click()
    If Not Dat_OriginesEthniques.Recordset.EOF Then
        stb_OriginesEthniques.Panels(4).Text = "SUPPRESSION"
        Load frm_OrigineEthniqueFiche
        frm_OrigineEthniqueFiche.Caption = "Suppression d'une origine ethnique"
        frm_OrigineEthniqueFiche.Txt_Code.Enabled = False
        frm_OrigineEthniqueFiche.Txt_Nom.Enabled = False
        
        frm_OrigineEthniqueFiche.Txt_Code = Dat_OriginesEthniques.Recordset("ORET_CODE")
        If Dat_OriginesEthniques.Recordset("ORET_LIBELLE") <> "" Then
            frm_OrigineEthniqueFiche.Txt_Nom = Dat_OriginesEthniques.Recordset("ORET_LIBELLE")
        Else
            frm_OrigineEthniqueFiche.Txt_Nom = ""
        End If
        frm_OrigineEthniqueFiche.Show vbModal
        
        stb_OriginesEthniques.Panels(4).Text = ""
    End If
End Sub

Private Sub cmd_Valider_Click()
    If Not Dat_OriginesEthniques.Recordset.EOF Then
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Dat_OriginesEthniques.DatabaseName = CheminBase
    Cmd_Actualiser_Click
End Sub

Private Sub grd_OriginesEthniques_HeadClick(ByVal ColIndex As Integer)

  Dim colonne As String
  Dim classem As String
  
  colonne = grd_OriginesEthniques.Columns(ColIndex).DataField
  Call Tri(Dat_OriginesEthniques, grd_OriginesEthniques, colonne)
      
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_OriginesEthniques.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_OriginesEthniques.Panels(4).Text = "Classement " & classem
    
End Sub

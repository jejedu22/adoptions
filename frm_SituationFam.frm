VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_SituationFam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Situations familiales"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_SituationFam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Dat_SF 
      Caption         =   "Dat_SF"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SITF"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSDBGrid.DBGrid grd_SF 
      Bindings        =   "frm_SituationFam.frx":08CA
      Height          =   5055
      Left            =   2760
      OleObjectBlob   =   "frm_SituationFam.frx":08DF
      TabIndex        =   9
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Actualiser 
         Default         =   -1  'True
         Height          =   375
         Left            =   7320
         Picture         =   "frm_SituationFam.frx":12C9
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
         Picture         =   "frm_SituationFam.frx":13CB
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
         Picture         =   "frm_SituationFam.frx":14CD
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Excel 
         Height          =   375
         Left            =   4560
         Picture         =   "frm_SituationFam.frx":15CF
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Export EXCEL"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "frm_SituationFam.frx":1759
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprime la situation familiale sélectionnée"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "frm_SituationFam.frx":185B
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Créer un nouvelle situation familiale"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "frm_SituationFam.frx":195D
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modifier les données de la situation familiale sélectionnée"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   9255
      Begin VB.TextBox Txt_SF_lib 
         Height          =   285
         Left            =   5520
         MaxLength       =   30
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox txt_SF_code 
         Height          =   285
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   7
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton cmd_SF_lib 
         Caption         =   "&Libellé"
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_SF_code 
         Caption         =   "&Code"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.StatusBar stb_SF 
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
            TextSave        =   "09:35"
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
Attribute VB_Name = "frm_SituationFam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Actualiser_Click()
    Dim sel_code As String
    Dim sel_lib As String
    
    If txt_SF_code <> "" Then sel_code = " AND SITF_CODE like """ & txt_SF_code & "*"""
    If Txt_SF_lib <> "" Then sel_lib = " AND SITF_DESIG like """ & Txt_SF_lib & "*"""
        
    Dat_SF.RecordSource = "SELECT * FROM SITF WHERE TRUE" + _
        sel_code + _
        sel_lib + _
        ClauseOrder
        
    Me.MousePointer = vbHourglass
    Dat_SF.Refresh
    Me.MousePointer = vbDefault
End Sub

Private Sub cmd_Creer_Click()
    stb_SF.Panels(4).Text = "CREATION"
    Load frm_SituationFamFiche
    frm_SituationFamFiche.Caption = "Création d'une institution"
    frm_SituationFamFiche.Txt_Code.Enabled = False
    frm_SituationFamFiche.Txt_Code = NouveauCode("SITF", "SITF_CODE")
    frm_SituationFamFiche.Txt_Nom = ""
    frm_SituationFamFiche.Show vbModal
  
    stb_SF.Panels(4).Text = ""
End Sub

Private Sub cmd_Excel_Click()

  ExportExcel (Dat_SF.RecordSource)
  
End Sub

Private Sub cmd_Modifier_Click()
    If Not Dat_SF.Recordset.EOF Then
        stb_SF.Panels(4).Text = "MODIFICATION"
        Load frm_SituationFamFiche
        frm_SituationFamFiche.Caption = "Modification d'une situation familiale"
        frm_SituationFamFiche.Txt_Code.Enabled = False
        frm_SituationFamFiche.Txt_Code = Dat_SF.Recordset("SITF_CODE")
        If Dat_SF.Recordset("SITF_DESIG") <> "" Then
            frm_SituationFamFiche.Txt_Nom = Dat_SF.Recordset("SITF_DESIG")
        Else
            frm_SituationFamFiche.Txt_Nom = ""
        End If
        frm_SituationFamFiche.Show vbModal
        
        stb_SF.Panels(4).Text = ""
    End If
End Sub

Private Sub cmd_Quitter_Click()
    Unload frm_SituationFam
End Sub

Private Sub cmd_SF_code_Click()
    txt_SF_code = ""
    txt_SF_code.SetFocus
End Sub

Private Sub cmd_SF_lib_Click()
    Txt_SF_lib = ""
    Txt_SF_lib.SetFocus
End Sub

Private Sub cmd_Supprimer_Click()
    If Not Dat_SF.Recordset.EOF Then
        stb_SF.Panels(4).Text = "SUPPRESSION"
        Load frm_SituationFamFiche
        frm_SituationFamFiche.Caption = "Suppression d'une situation familiale"
        frm_SituationFamFiche.Txt_Code.Enabled = False
        frm_SituationFamFiche.Txt_Nom.Enabled = False
        
        frm_SituationFamFiche.Txt_Code = Dat_SF.Recordset("SITF_CODE")
        If Dat_SF.Recordset("SITF_DESIG") <> "" Then
            frm_SituationFamFiche.Txt_Nom = Dat_SF.Recordset("SITF_DESIG")
        Else
            frm_SituationFamFiche.Txt_Nom = ""
        End If
        frm_SituationFamFiche.Show vbModal
        
        stb_SF.Panels(4).Text = ""
    End If
End Sub

Private Sub cmd_Valider_Click()
    If Not Dat_SF.Recordset.EOF Then
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Dat_SF.DatabaseName = CheminBase
    Cmd_Actualiser_Click
End Sub

Private Sub grd_SF_HeadClick(ByVal ColIndex As Integer)

  Dim colonne As String
  Dim classem As String
  
  colonne = grd_SF.Columns(ColIndex).DataField
  Call Tri(Dat_SF, grd_SF, colonne)
      
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_SF.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_SF.Panels(4).Text = "Classement " & classem
  
End Sub

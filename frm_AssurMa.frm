VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_AssurMa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Régimes d'assurance maladie"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_AssurMa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Dat_AssurMa 
      Caption         =   "Dat_AssuranceMaladie"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ASMA"
      Top             =   6000
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSDBGrid.DBGrid grd_AssurMa 
      Bindings        =   "frm_AssurMa.frx":08CA
      Height          =   5055
      Left            =   2280
      OleObjectBlob   =   "frm_AssurMa.frx":08E4
      TabIndex        =   9
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   9255
      Begin VB.CommandButton cmd_AssurMa_code 
         Caption         =   "&Code"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_AssurMa_nom 
         Caption         =   "&Libellé"
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txt_AssurMa_code 
         Height          =   285
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   7
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox txt_AssurMa_nom 
         Height          =   285
         Left            =   5520
         MaxLength       =   30
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "frm_AssurMa.frx":12D3
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modifier le régime d'assurance maladie sélectionné"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "frm_AssurMa.frx":13D5
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Créer un nouveau régime d'assurance maladie"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "frm_AssurMa.frx":14D7
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprimer le régime d'assurance maladie sélectionné"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmd_Excel 
         Height          =   375
         Left            =   4560
         Picture         =   "frm_AssurMa.frx":15D9
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Export EXCEL"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8760
         Picture         =   "frm_AssurMa.frx":1763
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8040
         Picture         =   "frm_AssurMa.frx":1865
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Valider"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmd_Actualiser 
         Default         =   -1  'True
         Height          =   375
         Left            =   7320
         Picture         =   "frm_AssurMa.frx":1967
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Rechercher d'après les critères"
         Top             =   120
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_ListeAssurMa 
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
            TextSave        =   "09:24"
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
Attribute VB_Name = "frm_AssurMa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Actualiser_Click()
    Dim sel_code As String
    Dim sel_nom As String
    
    If txt_AssurMa_code <> "" Then sel_code = " AND ASMA_CODE like """ & txt_AssurMa_code & "*"""
    If txt_AssurMa_nom <> "" Then sel_nom = " AND ASMA_DESIG like """ & txt_AssurMa_nom & "*"""
        
    Dat_AssurMa.RecordSource = "SELECT * FROM ASMA WHERE TRUE" + _
        sel_code + _
        sel_nom + _
        ClauseOrder
        
    Me.MousePointer = vbHourglass
    Dat_AssurMa.Refresh
    Me.MousePointer = vbDefault
End Sub

Private Sub cmd_AssurMa_code_Click()
    txt_AssurMa_code = ""
    txt_AssurMa_code.SetFocus
End Sub

Private Sub cmd_AssurMa_nom_Click()
    txt_AssurMa_nom = ""
    txt_AssurMa_nom.SetFocus
End Sub

Private Sub cmd_Creer_Click()
    stb_ListeAssurMa.Panels(4).Text = "CREATION"
    Load frm_AssurMaFiche
    frm_AssurMaFiche.Caption = "Création d'un régime d'assurance maladie"
    frm_AssurMaFiche.Txt_Code.Enabled = False
    frm_AssurMaFiche.Txt_Code = NouveauCode("ASMA", "ASMA_CODE")
    frm_AssurMaFiche.Txt_Nom = ""
    frm_AssurMaFiche.Show vbModal
  
    stb_ListeAssurMa.Panels(4).Text = ""
End Sub

Private Sub cmd_Excel_Click()

ExportExcel (Dat_AssurMa.RecordSource)

End Sub

Private Sub cmd_Modifier_Click()
    If Not Dat_AssurMa.Recordset.EOF Then
        stb_ListeAssurMa.Panels(4).Text = "MODIFICATION"
        Load frm_AssurMaFiche
        frm_AssurMaFiche.Caption = "Modification d'un régime d'assurance maladie"
        frm_AssurMaFiche.Txt_Code.Enabled = False
        frm_AssurMaFiche.Txt_Code = Dat_AssurMa.Recordset("ASMA_CODE")
        If Dat_AssurMa.Recordset("ASMA_DESIG") <> "" Then
            frm_AssurMaFiche.Txt_Nom = Dat_AssurMa.Recordset("ASMA_DESIG")
        Else
            frm_AssurMaFiche.Txt_Nom = ""
        End If
        frm_AssurMaFiche.Show vbModal
        
        stb_ListeAssurMa.Panels(4).Text = ""
    End If
End Sub

Private Sub cmd_Quitter_Click()
    Unload Me
End Sub

Private Sub cmd_Supprimer_Click()
    If Not Dat_AssurMa.Recordset.EOF Then
        stb_ListeAssurMa.Panels(4).Text = "SUPPRESSION"
        Load frm_AssurMaFiche
        frm_AssurMaFiche.Caption = "Suppression d'un régime d'assurance maladie"
        frm_AssurMaFiche.Txt_Code.Enabled = False
        frm_AssurMaFiche.Txt_Nom.Enabled = False
        
        frm_AssurMaFiche.Txt_Code = Dat_AssurMa.Recordset("ASMA_CODE")
        If Dat_AssurMa.Recordset("ASMA_DESIG") <> "" Then
            frm_AssurMaFiche.Txt_Nom = Dat_AssurMa.Recordset("ASMA_DESIG")
        Else
            frm_AssurMaFiche.Txt_Nom = ""
        End If
        frm_AssurMaFiche.Show vbModal
        
        stb_ListeAssurMa.Panels(4).Text = ""
    End If
End Sub

Private Sub cmd_Valider_Click()
    If Not Dat_AssurMa.Recordset.EOF Then
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Dat_AssurMa.DatabaseName = CheminBase
    Cmd_Actualiser_Click
End Sub

Private Sub grd_AssurMa_HeadClick(ByVal ColIndex As Integer)

Dim colonne As String
Dim classem As String
   
  colonne = grd_AssurMa.Columns(ColIndex).DataField
  Call Tri(Dat_AssurMa, grd_AssurMa, colonne)
      
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_AssurMa.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_ListeAssurMa.Panels(4).Text = "Classement " & classem
    
End Sub


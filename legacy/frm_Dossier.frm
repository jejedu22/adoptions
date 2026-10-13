VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Dossier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recherche d'un dossier dans la liste"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_Dossier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txt_CP 
      Height          =   285
      Left            =   2400
      MaxLength       =   5
      TabIndex        =   6
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Data Dat_Dossiers 
      Caption         =   "Dat_Dossiers"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Actualiser 
         Default         =   -1  'True
         Height          =   375
         Left            =   7320
         Picture         =   "frm_Dossier.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Rechercher d'après les critères"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8040
         Picture         =   "frm_Dossier.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8760
         Picture         =   "frm_Dossier.frx":0ACE
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Excel 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   4560
         Picture         =   "frm_Dossier.frx":0BD0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Export EXCEL"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_ListeDossiers 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   6075
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
            TextSave        =   "09/12/99"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "17:16"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10557
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   9255
      Begin VB.TextBox Txt_PrenomDemandeur 
         Height          =   285
         Left            =   5640
         MaxLength       =   20
         TabIndex        =   21
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Txt_NomDemandeur 
         Height          =   285
         Left            =   5640
         MaxLength       =   20
         TabIndex        =   20
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Cmd_PrenomDemandeur 
         Caption         =   "&Prénom demandeur"
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Cmd_NomDemandeur 
         Caption         =   "No&m demandeur"
         Height          =   255
         Left            =   3960
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmd_numero 
         Caption         =   "&N° de  dossier"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_reinst 
         Caption         =   "&Réinstruction"
         Height          =   255
         Left            =   3960
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txt_code 
         Height          =   285
         Left            =   2280
         MaxLength       =   8
         TabIndex        =   4
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton cmd_CP 
         Caption         =   "&Code Postal"
         Height          =   255
         Left            =   720
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmd_Ville 
         Caption         =   "&Ville"
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmd_Tel 
         Caption         =   "N° &Téléphone"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Txt_Tel 
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   8
         Top             =   960
         Width           =   1395
      End
      Begin VB.TextBox Txt_Ville 
         Height          =   285
         Left            =   5640
         MaxLength       =   30
         TabIndex        =   7
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox Cmb_reins 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin MSDBGrid.DBGrid grd_Dossiers 
      Bindings        =   "frm_Dossier.frx":0D5A
      Height          =   3615
      Left            =   120
      OleObjectBlob   =   "frm_Dossier.frx":0D75
      TabIndex        =   9
      Top             =   2400
      Width           =   9255
   End
End
Attribute VB_Name = "frm_Dossier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Actualiser_Click()
    Dim sel_code As String
    Dim sel_reinst As String
    Dim sel_cp As String
    Dim sel_ville As String
    Dim sel_tel As String
    Dim sel_nom As String
    Dim sel_prenom As String
    Dim reinsSEARCH As String
        
    If Cmb_reins.ListIndex <> 0 Then
        reinstSEARCH = Cmb_reins.Text
    Else
        reinstSEARCH = ""
    End If
        
    If Txt_Code <> "" Then sel_code = " AND DOSS.DOSS_NUM like """ & Txt_Code & "*"""
    If Txt_CP <> "" Then sel_cp = " AND DOSS_CP like """ & Txt_CP & "*"""
    If reinstSEARCH <> "" Then sel_reinst = " AND DOSS_REINS like """ & reinstSEARCH & "*"""
    If Txt_Ville <> "" Then sel_ville = " AND DOSS_VILLE like """ & Txt_Ville & "*"""
    If Txt_Tel <> "" Then sel_tel = " AND DOSS_NUMTEL like """ & Txt_Tel & "*"""
    If Txt_NomDemandeur <> "" Then sel_nom = " AND INDI_NOM like """ & Txt_NomDemandeur & "*"""
    If Txt_PrenomDemandeur <> "" Then sel_prenom = " AND INDI_PREN like """ & Txt_PrenomDemandeur & "*"""
    Dat_Dossiers.RecordSource = "SELECT DOSS.DOSS_NUM,DOSS.DOSS_REINS,DOSS.DOSS_CP,DOSS.DOSS_VILLE,INDI.INDI_NOM,INDI.INDI_PREN,DOSS.DOSS_NUMTEL FROM DOSS,INDI WHERE TRUE" + _
        sel_code + _
        sel_reinst + _
        sel_cp + _
        sel_ville + _
        sel_tel + _
        sel_nom + _
        sel_prenom + _
        " AND INDI.DOSS_NUM = DOSS.DOSS_NUM AND INDI.INDI_CDFR= ""DE"" " + _
        ClauseOrder
        
    Me.MousePointer = vbHourglass
    Dat_Dossiers.Refresh
    Me.MousePointer = vbDefault
End Sub

Private Sub cmd_CP_Click()
    Txt_CP = ""
    Txt_CP.SetFocus
End Sub

Private Sub cmd_Excel_Click()

  ExportExcel (Dat_Dossiers.RecordSource)

End Sub

Private Sub Cmd_NomDemandeur_Click()
    Txt_NomDemandeur = ""
    Txt_NomDemandeur.SetFocus
End Sub

Private Sub cmd_numero_Click()
    Txt_Code = ""
    Txt_Code.SetFocus
End Sub

Private Sub Cmd_PrenomDemandeur_Click()
    Txt_PrenomDemandeur = ""
    Txt_PrenomDemandeur.SetFocus
End Sub

Private Sub cmd_Quitter_Click()
    Unload Me
End Sub

Private Sub cmd_reinst_Click()
    Cmb_reins.ListIndex = 0
    Cmb_reins.SetFocus
End Sub

Private Sub cmd_Tel_Click()
    Txt_Tel = ""
    Txt_Tel.SetFocus
End Sub

Private Sub cmd_Valider_Click()

  If Not Dat_Dossiers.Recordset.EOF Then
    frm_DossierGestion.Cmb_NumDossier.Text = Dat_Dossiers.Recordset("DOSS_NUM")
  End If
  Unload Me
  
End Sub

Private Sub cmd_Ville_Click()
    Txt_Ville = ""
    Txt_Ville.SetFocus
End Sub

Private Sub Form_Load()
    Dat_Dossiers.DatabaseName = CheminBase
    Dat_Dossiers.RecordSource = "SELECT DOSS.DOSS_NUM,DOSS.DOSS_REINS,DOSS.DOSS_CP,DOSS.DOSS_VILLE,INDI.INDI_NOM,INDI.INDI_PREN,DOSS.DOSS_NUMTEL " + _
        "FROM DOSS,INDI WHERE INDI.DOSS_NUM = DOSS.DOSS_NUM AND INDI.INDI_CDFR = ""DE"" "
    Dat_Dossiers.Refresh
    Cmd_Actualiser_Click
    ListeReinstruction Cmb_reins
    Cmb_reins.ListIndex = 0
End Sub

Private Sub grd_Dossiers_HeadClick(ByVal ColIndex As Integer)
    Dim colonne As String
    
    colonne = grd_Dossiers.Columns(ColIndex).DataField
    If ClauseOrder = " ORDER BY " + colonne + " ASC" Then
        ClauseOrder = " ORDER BY " + colonne + " DESC"
    Else
        ClauseOrder = " ORDER BY " + colonne + " ASC"
    End If
    If InStr(Dat_Dossiers.RecordSource, " ORDER BY ") <> 0 Then
        Dat_Dossiers.RecordSource = Left$(Dat_Dossiers.RecordSource, InStr(Dat_Dossiers.RecordSource, "ORDER BY") - 1) + ClauseOrder
    Else
        Dat_Dossiers.RecordSource = Dat_Dossiers.RecordSource + ClauseOrder
    End If
    Dat_Dossiers.Refresh
End Sub

Public Sub ListeReinstruction(Boite As ComboBox)
'Permet de choisir un dossier en réinstruction ou non
'grâce à la Combo
    With frm_Dossier.Dat_Dossiers
        Boite.AddItem ("<Tous>")
        Boite.AddItem ("O")
        Boite.AddItem ("N")
    End With
End Sub



VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Tribunaux 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tribunaux judiciaires"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_Tribunaux.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Dat_Trib 
      Caption         =   "Dat_Trib"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4155
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   9255
      Begin VB.TextBox txt_ville 
         Height          =   285
         Left            =   5520
         MaxLength       =   5
         TabIndex        =   10
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txt_cp 
         Height          =   285
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   9
         Top             =   600
         Width           =   1395
      End
      Begin VB.CommandButton cmd_ville 
         Caption         =   "&Ville"
         Height          =   255
         Left            =   3960
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmd_cp 
         Caption         =   "Code &Postal"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txt_nom 
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
      Begin VB.CommandButton cmd_nom 
         Caption         =   "&Libellé"
         Height          =   255
         Left            =   3960
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_code 
         Caption         =   "&Code"
         Height          =   255
         Left            =   720
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Actualiser 
         Default         =   -1  'True
         Height          =   375
         Left            =   7320
         Picture         =   "frm_Tribunaux.frx":0442
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
         Picture         =   "frm_Tribunaux.frx":0544
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
         Picture         =   "frm_Tribunaux.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Excel 
         Height          =   375
         Left            =   4560
         Picture         =   "frm_Tribunaux.frx":0748
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Export EXCEL"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "frm_Tribunaux.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprimer le tribunal sélectionné"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "frm_Tribunaux.frx":09D4
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Créer un nouveau tribunal"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "frm_Tribunaux.frx":0AD6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modifier les données du tribunal sélectionné"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_ListeTribunaux 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
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
            TextSave        =   "09:36"
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
   Begin MSDBGrid.DBGrid grd_Tribunaux 
      Bindings        =   "frm_Tribunaux.frx":0BD8
      Height          =   4575
      Left            =   120
      OleObjectBlob   =   "frm_Tribunaux.frx":0BEF
      TabIndex        =   11
      Top             =   1920
      Width           =   9255
   End
End
Attribute VB_Name = "frm_Tribunaux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Actualiser_Click()

  Dim sel_code As String
  Dim sel_nom As String
  Dim sel_cp As String
  Dim sel_ville As String
      
  If Txt_Code <> "" Then sel_code = " AND JUGE_NUM like """ & Txt_Code & "*"""
  If Txt_Nom <> "" Then sel_nom = " AND JUGE_NOM like """ & Txt_Nom & "*"""
      
  Dat_Trib.RecordSource = "SELECT * FROM JUGE WHERE TRUE" + _
      sel_code + sel_nom + _
      sel_cp + _
      sel_ville + _
      ClauseOrder
      
  Me.MousePointer = vbHourglass
  Dat_Trib.Refresh
  Me.MousePointer = vbDefault
  
End Sub

Private Sub cmd_code_Click()
    Txt_Code = ""
    Txt_Code.SetFocus
End Sub

Private Sub cmd_corres_Click()
    txt_corres = ""
    txt_corres.SetFocus
End Sub

Private Sub cmd_CP_Click()
    Txt_CP = ""
    Txt_CP.SetFocus
End Sub

Private Sub cmd_Creer_Click()
    stb_ListeTribunaux.Panels(4).Text = "CREATION"
    Load frm_TribunauxFiche
    frm_TribunauxFiche.Caption = "Création d'un tribunal"
    frm_TribunauxFiche.Txt_Code.Enabled = False
    frm_TribunauxFiche.Txt_Code = NouveauCode("JUGE", "JUGE_NUM")
    frm_TribunauxFiche.Txt_Nom = ""
    frm_TribunauxFiche.Txt_adr1 = ""
    frm_TribunauxFiche.Txt_adr2 = ""
    frm_TribunauxFiche.Txt_CP = ""
    frm_TribunauxFiche.Txt_Ville = ""
    frm_TribunauxFiche.txt_corres = ""
    frm_TribunauxFiche.Show vbModal
  
    stb_ListeTribunaux.Panels(4).Text = ""
End Sub

Private Sub cmd_Excel_Click()
  
  ExportExcel (Dat_Trib.RecordSource)
  
End Sub

Private Sub cmd_Modifier_Click()
    If Not Dat_Trib.Recordset.EOF Then
        stb_ListeTribunaux.Panels(4).Text = "MODIFICATION"
        Load frm_TribunauxFiche
        frm_TribunauxFiche.Caption = "Modification d'un tribunal"
        frm_TribunauxFiche.Txt_Code.Enabled = False
        frm_TribunauxFiche.Txt_Code = Dat_Trib.Recordset("JUGE_NUM")
        If Dat_Trib.Recordset("JUGE_NOM") <> "" Then
            frm_TribunauxFiche.Txt_Nom = Dat_Trib.Recordset("JUGE_NOM")
        Else
            frm_TribunauxFiche.Txt_Nom = ""
        End If
        If Dat_Trib.Recordset("JUGE_ADR1") <> "" Then
            frm_TribunauxFiche.Txt_adr1 = Dat_Trib.Recordset("JUGE_ADR1")
        Else
            frm_TribunauxFiche.Txt_adr1 = ""
        End If
        If Dat_Trib.Recordset("JUGE_ADR2") <> "" Then
            frm_TribunauxFiche.Txt_adr2 = Dat_Trib.Recordset("JUGE_ADR2")
        Else
            frm_TribunauxFiche.Txt_adr2 = ""
        End If
        If Dat_Trib.Recordset("JUGE_CPTGI") <> "" Then
            frm_TribunauxFiche.Txt_CP = Dat_Trib.Recordset("JUGE_CPTGI")
        Else
            frm_TribunauxFiche.Txt_CP = ""
        End If
        If Dat_Trib.Recordset("JUGE_VILLE") <> "" Then
            frm_TribunauxFiche.Txt_Ville = Dat_Trib.Recordset("JUGE_VILLE")
        Else
            frm_TribunauxFiche.Txt_Ville = ""
        End If
        If Dat_Trib.Recordset("JUGE_NOMC") <> "" Then
            frm_TribunauxFiche.txt_corres = Dat_Trib.Recordset("JUGE_NOMC")
        Else
            frm_TribunauxFiche.txt_corres = ""
        End If
        
        frm_TribunauxFiche.Show vbModal
        
        stb_ListeTribunaux.Panels(4).Text = ""
    End If
End Sub

Private Sub cmd_nom_Click()
    Txt_Nom = ""
    Txt_Nom.SetFocus
End Sub

Private Sub cmd_Quitter_Click()
    Unload Me
End Sub

Private Sub cmd_Supprimer_Click()
    If Not Dat_Trib.Recordset.EOF Then
        stb_ListeTribunaux.Panels(4).Text = "SUPPRESSION"
        Load frm_TribunauxFiche
        frm_TribunauxFiche.Caption = "Suppression d'un tribunal"
        frm_TribunauxFiche.Txt_adr1.Enabled = False
        frm_TribunauxFiche.Txt_adr2.Enabled = False
        frm_TribunauxFiche.Txt_Code.Enabled = False
        frm_TribunauxFiche.txt_corres.Enabled = False
        frm_TribunauxFiche.Txt_CP.Enabled = False
        frm_TribunauxFiche.Txt_Nom.Enabled = False
        frm_TribunauxFiche.Txt_Ville.Enabled = False
        
        frm_TribunauxFiche.Txt_Code = Dat_Trib.Recordset("JUGE_NUM")
        If Dat_Trib.Recordset("JUGE_NOM") <> "" Then
            frm_TribunauxFiche.Txt_Nom = Dat_Trib.Recordset("JUGE_NOM")
        Else
            frm_TribunauxFiche.Txt_Nom = ""
        End If
        If Dat_Trib.Recordset("JUGE_ADR1") <> "" Then
            frm_TribunauxFiche.Txt_adr1 = Dat_Trib.Recordset("JUGE_ADR1")
        Else
            frm_TribunauxFiche.Txt_adr1 = ""
        End If
        If Dat_Trib.Recordset("JUGE_ADR2") <> "" Then
            frm_TribunauxFiche.Txt_adr2 = Dat_Trib.Recordset("JUGE_ADR2")
        Else
            frm_TribunauxFiche.Txt_adr2 = ""
        End If
        If Dat_Trib.Recordset("JUGE_CPTGI") <> "" Then
            frm_TribunauxFiche.Txt_CP = Dat_Trib.Recordset("JUGE_CPTGI")
        Else
            frm_TribunauxFiche.Txt_CP = ""
        End If
        If Dat_Trib.Recordset("JUGE_VILLE") <> "" Then
            frm_TribunauxFiche.Txt_Ville = Dat_Trib.Recordset("JUGE_VILLE")
        Else
            frm_TribunauxFiche.Txt_Ville = ""
        End If
        If Dat_Trib.Recordset("JUGE_NOMC") <> "" Then
            frm_TribunauxFiche.txt_corres = Dat_Trib.Recordset("JUGE_NOMC")
        Else
            frm_TribunauxFiche.txt_corres = ""
        End If
                
        frm_TribunauxFiche.Show vbModal
        
        stb_ListeTribunaux.Panels(4).Text = ""
    End If
End Sub

Private Sub cmd_Valider_Click()
    If Not Dat_Trib.Recordset.EOF Then
        Me.Hide
    End If
End Sub

Private Sub cmd_Ville_Click()
    Txt_Ville = ""
    Txt_Ville.SetFocus
End Sub

Private Sub Form_Load()

  With Dat_Trib
    .DatabaseName = CheminBase
    .RecordSource = "select * from juge"
    .Refresh
  End With
  
End Sub

Private Sub grd_Tribunaux_HeadClick(ByVal ColIndex As Integer)

  Dim colonne As String
  Dim classem As String
  
  colonne = grd_Tribunaux.Columns(ColIndex).DataField
  Call Tri(Dat_Trib, grd_Tribunaux, colonne)
      
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_Trib.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_ListeTribunaux.Panels(4).Text = "Classement " & classem
  
End Sub

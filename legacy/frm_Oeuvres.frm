VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Oeuvres 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oeuvres de placement"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_Oeuvres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Dat_oeuvre 
      Caption         =   "Dat_Oeuvre"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "OEPL"
      Top             =   5760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid grd_oeuvres 
      Bindings        =   "frm_Oeuvres.frx":0442
      Height          =   4335
      Left            =   120
      OleObjectBlob   =   "frm_Oeuvres.frx":045B
      TabIndex        =   13
      Top             =   2160
      Width           =   9255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      Height          =   1455
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Width           =   9255
      Begin VB.CommandButton cmd_oepl_nomres 
         Caption         =   "&Responsable"
         Height          =   255
         Left            =   3960
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txt_oepl_tel 
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txt_oepl_nomres 
         Height          =   285
         Left            =   5520
         TabIndex        =   12
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton cmd_oepl_tel 
         Caption         =   "Télép&hone"
         Height          =   255
         Left            =   720
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txt_oepl_nom 
         Height          =   285
         Left            =   5520
         MaxLength       =   50
         TabIndex        =   8
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txt_oepl_code 
         Height          =   285
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   7
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton cmd_oepl_nom 
         Caption         =   "N&om"
         Height          =   255
         Left            =   3960
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_oepl_code 
         Caption         =   "&N° Oeuvre"
         Height          =   255
         Left            =   720
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_oepl_cp 
         Caption         =   "Code &Postal"
         Height          =   255
         Left            =   720
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txt_oepl_cp 
         Height          =   285
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   9
         Top             =   600
         Width           =   675
      End
      Begin VB.CommandButton cmd_oepl_ville 
         Caption         =   "&Ville"
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txt_oepl_ville 
         Height          =   285
         Left            =   5520
         MaxLength       =   30
         TabIndex        =   10
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Actualiser 
         Default         =   -1  'True
         Height          =   375
         Left            =   7320
         Picture         =   "frm_Oeuvres.frx":183E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "REchercher d'après les critères"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         Height          =   375
         Left            =   8040
         Picture         =   "frm_Oeuvres.frx":1940
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frm_Oeuvres.frx":1A42
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Excel 
         Height          =   375
         Left            =   4560
         Picture         =   "frm_Oeuvres.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Export EXCEL"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "frm_Oeuvres.frx":1CCE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprime l'oeuvre de placement sélectionnée"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "frm_Oeuvres.frx":1DD0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Créer une nouvelle oeuvre de placement"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "frm_Oeuvres.frx":1ED2
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modifier les données de l'oeuvre de placement sélectionnée"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_ListeOeuvres 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   18
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
            Object.Width           =   10557
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_Oeuvres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Actualiser_Click()
    Dim sel_code As String
    Dim sel_nom As String
    Dim sel_cp As String
    Dim sel_ville As String
    Dim sel_tel As String
    Dim sel_nomres As String
    
    If txt_oepl_code <> "" Then sel_code = " AND OEPL_CODE like """ & txt_oepl_code & "*"""
    If txt_oepl_nom <> "" Then sel_nom = " AND OEPL_DESIGN like """ & txt_oepl_nom & "*"""
    If txt_oepl_cp <> "" Then sel_cp = " AND OEPL_CP like """ & txt_oepl_cp & "*"""
    If txt_oepl_ville <> "" Then sel_ville = " AND OEPL_VILLE like """ & txt_oepl_ville & "*"""
    If txt_oepl_tel <> "" Then sel_tel = " AND OEPL_TEL like """ & txt_oepl_tel & "*"""
    If txt_oepl_nomres <> "" Then sel_nomres = " AND OEPL_NOMRES like """ & txt_oepl_nomres & "*"""
    
    Dat_oeuvre.RecordSource = "SELECT * FROM OEPL WHERE TRUE" + _
        sel_code + _
        sel_nom + _
        sel_cp + _
        sel_ville + _
        sel_tel + _
        sel_nomres + _
        ClauseOrder
        
    Me.MousePointer = vbHourglass
    Dat_oeuvre.Refresh
    Me.MousePointer = vbDefault
End Sub

Private Sub cmd_Creer_Click()
    stb_ListeOeuvres.Panels(4).Text = "CREATION"
    Load frm_OeuvreFiche
    frm_OeuvreFiche.Caption = "Création d'une oeuvre de placement"
    frm_OeuvreFiche.Txt_Code.Enabled = False
    frm_OeuvreFiche.Txt_Code = NouveauCode("OEPL", "OEPL_CODE")
    frm_OeuvreFiche.Txt_Nom = ""
    frm_OeuvreFiche.Txt_Adr = ""
    frm_OeuvreFiche.Txt_CP = ""
    frm_OeuvreFiche.Txt_Ville = ""
    frm_OeuvreFiche.Txt_Tel = ""
    frm_OeuvreFiche.Txt_nomres = ""
    frm_OeuvreFiche.Txt_rem = ""
    frm_OeuvreFiche.Show vbModal
  
    stb_ListeOeuvres.Panels(4).Text = ""
End Sub

Private Sub cmd_Excel_Click()

  ExportExcel (Dat_oeuvre.RecordSource)

End Sub

Private Sub cmd_Modifier_Click()

  If Not Dat_oeuvre.Recordset.EOF Then
    stb_ListeOeuvres.Panels(4).Text = "MODIFICATION"
    Load frm_OeuvreFiche
    frm_OeuvreFiche.Caption = "Modification d'une oeuvre de placement"
    frm_OeuvreFiche.Txt_Code.Enabled = False
    frm_OeuvreFiche.Txt_Code = Dat_oeuvre.Recordset("OEPL_CODE")
    If Dat_oeuvre.Recordset("OEPL_DESIGN") <> "" Then
        frm_OeuvreFiche.Txt_Nom = Dat_oeuvre.Recordset("OEPL_DESIGN")
    Else
        frm_OeuvreFiche.Txt_Nom = ""
    End If
    If Dat_oeuvre.Recordset("OEPL_ADR") <> "" Then
        frm_OeuvreFiche.Txt_Adr = Dat_oeuvre.Recordset("OEPL_ADR")
    Else
        frm_OeuvreFiche.Txt_Adr = ""
    End If
    If Dat_oeuvre.Recordset("OEPL_CP") <> "" Then
        frm_OeuvreFiche.Txt_CP = Dat_oeuvre.Recordset("OEPL_CP")
    Else
        frm_OeuvreFiche.Txt_CP = ""
    End If
    If Dat_oeuvre.Recordset("OEPL_VILLE") <> "" Then
        frm_OeuvreFiche.Txt_Ville = Dat_oeuvre.Recordset("OEPL_VILLE")
    Else
        frm_OeuvreFiche.Txt_Ville = ""
    End If
    If Dat_oeuvre.Recordset("OEPL_TEL") <> "" Then
        frm_OeuvreFiche.Txt_Tel = Dat_oeuvre.Recordset("OEPL_TEL")
    Else
        frm_OeuvreFiche.Txt_Tel = ""
    End If
    If Dat_oeuvre.Recordset("OEPL_NOMRES") <> "" Then
        frm_OeuvreFiche.Txt_nomres = Dat_oeuvre.Recordset("OEPL_NOMRES")
    Else
        frm_OeuvreFiche.Txt_nomres = ""
    End If
    If Dat_oeuvre.Recordset("OEPL_REM") <> "" Then
        frm_OeuvreFiche.Txt_rem = Dat_oeuvre.Recordset("OEPL_REM")
    Else
        frm_OeuvreFiche.Txt_rem = ""
    End If
    If Dat_oeuvre.Recordset("OEPL_EMAIL") <> "" Then
        frm_OeuvreFiche.Txt_EMail = Dat_oeuvre.Recordset("OEPL_EMAIL")
    Else
        frm_OeuvreFiche.Txt_EMail = ""
    End If
    
    frm_OeuvreFiche.Show vbModal
    
    stb_ListeOeuvres.Panels(4).Text = ""
  End If
  
End Sub

Private Sub cmd_oepl_code_Click()
    txt_oepl_code = ""
    txt_oepl_code.SetFocus
End Sub

Private Sub cmd_oepl_cp_Click()
    txt_oepl_cp = ""
    txt_oepl_cp.SetFocus
End Sub

Private Sub cmd_oepl_nom_Click()
    txt_oepl_nom = ""
    txt_oepl_nom.SetFocus
End Sub

Private Sub cmd_oepl_nomres_Click()
    txt_oepl_nomres = ""
    txt_oepl_nomres.SetFocus
End Sub

Private Sub cmd_oepl_tel_Click()
    txt_oepl_tel = ""
    txt_oepl_tel.SetFocus
End Sub

Private Sub cmd_oepl_ville_Click()
    txt_oepl_ville = ""
    txt_oepl_ville.SetFocus
End Sub

Private Sub cmd_Quitter_Click()
    Unload frm_Oeuvres
End Sub

Private Sub cmd_Supprimer_Click()
    If Not Dat_oeuvre.Recordset.EOF Then
        stb_ListeOeuvres.Panels(4).Text = "SUPPRESSION"
        Load frm_OeuvreFiche
        frm_OeuvreFiche.Caption = "Suppression d'une oeuvre de placement"
        frm_OeuvreFiche.Fra_oeuvres.Enabled = False

        frm_OeuvreFiche.Txt_Code = Dat_oeuvre.Recordset("OEPL_CODE")
        If Dat_oeuvre.Recordset("OEPL_DESIGN") <> "" Then
            frm_OeuvreFiche.Txt_Nom = Dat_oeuvre.Recordset("OEPL_DESIGN")
        Else
            frm_OeuvreFiche.Txt_Nom = ""
        End If
        If Dat_oeuvre.Recordset("OEPL_ADR") <> "" Then
            frm_OeuvreFiche.Txt_Adr = Dat_oeuvre.Recordset("OEPL_ADR")
        Else
            frm_OeuvreFiche.Txt_Adr = ""
        End If
        If Dat_oeuvre.Recordset("OEPL_CP") <> "" Then
            frm_OeuvreFiche.Txt_CP = Dat_oeuvre.Recordset("OEPL_CP")
        Else
            frm_OeuvreFiche.Txt_CP = ""
        End If
        If Dat_oeuvre.Recordset("OEPL_VILLE") <> "" Then
            frm_OeuvreFiche.Txt_Ville = Dat_oeuvre.Recordset("OEPL_VILLE")
        Else
            frm_OeuvreFiche.Txt_Ville = ""
        End If
        If Dat_oeuvre.Recordset("OEPL_TEL") <> "" Then
            frm_OeuvreFiche.Txt_Tel = Dat_oeuvre.Recordset("OEPL_TEL")
        Else
            frm_OeuvreFiche.Txt_Tel = ""
        End If
        If Dat_oeuvre.Recordset("OEPL_NOMRES") <> "" Then
            frm_OeuvreFiche.Txt_nomres = Dat_oeuvre.Recordset("OEPL_NOMRES")
        Else
            frm_OeuvreFiche.Txt_nomres = ""
        End If
        If Dat_oeuvre.Recordset("OEPL_REM") <> "" Then
            frm_OeuvreFiche.Txt_rem = Dat_oeuvre.Recordset("OEPL_REM")
        Else
            frm_OeuvreFiche.Txt_rem = ""
        End If
        If Dat_oeuvre.Recordset("OEPL_EMAIL") <> "" Then
            frm_OeuvreFiche.Txt_EMail = Dat_oeuvre.Recordset("OEPL_EMAIL")
        Else
            frm_OeuvreFiche.Txt_EMail = ""
        End If
        
        frm_OeuvreFiche.Show vbModal
        
        stb_ListeOeuvres.Panels(4).Text = ""
    End If
End Sub

Private Sub cmd_Valider_Click()
    If Not Dat_oeuvre.Recordset.EOF Then
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Dat_oeuvre.DatabaseName = CheminBase
    Cmd_Actualiser_Click
End Sub

Private Sub grd_oeuvres_HeadClick(ByVal ColIndex As Integer)

  Dim colonne As String
  Dim classem As String
  
  colonne = grd_oeuvres.Columns(ColIndex).DataField
  Call Tri(Dat_oeuvre, grd_oeuvres, colonne)
      
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_oeuvre.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_ListeOeuvres.Panels(4).Text = "Classement " & classem
    
End Sub

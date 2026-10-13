VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_TypeRecoursListe 
   Caption         =   "Types de recours contentieux Liste"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   ControlBox      =   0   'False
   Icon            =   "Frm_TypeRecoursListe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Dat_maxcode 
      Caption         =   "Dat_maxcode"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Dat_tyre 
      Caption         =   "Dat_tyre"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSDBGrid.DBGrid DBG_TypeRecours 
      Bindings        =   "Frm_TypeRecoursListe.frx":08CA
      Height          =   2775
      Left            =   1200
      OleObjectBlob   =   "Frm_TypeRecoursListe.frx":08E1
      TabIndex        =   7
      Top             =   960
      Width           =   5175
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   6840
         Picture         =   "Frm_TypeRecoursListe.frx":12D4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Excel 
         Height          =   375
         Left            =   3960
         Picture         =   "Frm_TypeRecoursListe.frx":13D6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Export EXCEL"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "Frm_TypeRecoursListe.frx":1560
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprimer le type de recours contentieux sélectionné"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "Frm_TypeRecoursListe.frx":1662
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Créer un nouveau type de recours contencieux"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "Frm_TypeRecoursListe.frx":1764
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Modifier le type de recours contentieux sélectionné"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_Types 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4215
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "26/06/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15:46"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5900
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Frm_TypeRecoursListe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Valider_Click()

End Sub


Private Sub cmd_Creer_Click()

Dim code As Integer

  stb_Types.Panels(4).Text = "CREATION"
  Dat_maxcode.RecordSource = "select max(tyre_num) from tyre"
  Dat_maxcode.Refresh
  Load Frm_TypeRecoursFiche
  
  If Dat_tyre.Recordset.EOF Then
    code = 1
  Else
    code = Dat_maxcode.Recordset(0) + 1
  End If
  
  Frm_TypeRecoursFiche.Txt_Code = code
  Frm_TypeRecoursFiche.Txt_lib = ""
  Frm_TypeRecoursFiche.Show vbModal
  stb_Types.Panels(4).Text = ""
  
End Sub

Private Sub cmd_Excel_Click()
  ExportExcel (Dat_tyre.RecordSource)
End Sub


Private Sub cmd_Modifier_Click()

If Not Dat_tyre.Recordset.EOF Then
  stb_Types.Panels(4) = "MODIFICATION"
  Load Frm_TypeRecoursFiche
  Frm_TypeRecoursFiche.Txt_Code = Dat_tyre.Recordset("TYRE_NUM")
  Frm_TypeRecoursFiche.Txt_lib = Dat_tyre.Recordset("TYRE_NOM")
  Frm_TypeRecoursFiche.Show vbModal
  stb_Types.Panels(4) = ""
End If

End Sub


Private Sub cmd_Quitter_Click()
  Unload Me
End Sub


Private Sub cmd_Supprimer_Click()

If Not Dat_tyre.Recordset.EOF Then
  stb_Types.Panels(4) = "SUPPRESSION"
  Load Frm_TypeRecoursFiche
  Frm_TypeRecoursFiche.Txt_Code = Dat_tyre.Recordset("TYRE_NUM")
  Frm_TypeRecoursFiche.Txt_lib = Dat_tyre.Recordset("TYRE_NOM")
  Frm_TypeRecoursFiche.Txt_lib.Enabled = False
  Frm_TypeRecoursFiche.Show vbModal
  stb_Types.Panels(4) = ""
End If

End Sub

Private Sub DBG_TypeRecours_HeadClick(ByVal ColIndex As Integer)

    Dim colonne As String
    Dim classem As String
    
    colonne = DBG_TypeRecours.Columns(ColIndex).DataField
    Call Tri(Dat_tyre, DBG_TypeRecours, colonne)
        
   
    'On affiche dans la statusbar l'ordre de tri
    If InStr(Dat_tyre.RecordSource, "ASC") <> 0 Then
        classem = "Croissant"
    Else
        classem = "Décroissant"
    End If
           
    stb_Types.Panels(4).Text = "Classement " & classem

End Sub


Private Sub Form_Load()
  
  Dat_tyre.DatabaseName = CheminBase
  Dat_maxcode.DatabaseName = CheminBase
  Dat_tyre.RecordSource = "select * from tyre order by tyre_num"
  Dat_tyre.Refresh
  
End Sub



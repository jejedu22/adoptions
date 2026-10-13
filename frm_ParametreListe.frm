VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_ParametreListe 
   Caption         =   "Paramètres de l'échéancier"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data dat_para 
      Caption         =   "dat_para"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_ParametreListe.frx":0000
      Height          =   3255
      Left            =   840
      OleObjectBlob   =   "frm_ParametreListe.frx":0017
      TabIndex        =   7
      Top             =   960
      Width           =   7575
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "frm_ParametreListe.frx":0BC2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Modifier le paramètre sélectionné"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "frm_ParametreListe.frx":0CC4
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Créer un nouveau paramètre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "frm_ParametreListe.frx":0DC6
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprimer le paramètre sélectionné"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Excel 
         Height          =   375
         Left            =   4560
         Picture         =   "frm_ParametreListe.frx":0EC8
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Export EXCEL"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8760
         Picture         =   "frm_ParametreListe.frx":1052
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_ParametreListe 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4845
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
            TextSave        =   "09:31"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9075
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_ParametreListe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Creer_Click()

  stb_ParametreListe.Panels(3) = "CREATION"
  Load frm_ParametreFiche
  If dat_para.Recordset.RecordCount = 0 Then
    frm_ParametreFiche.Txt_Code = 1
  Else
    frm_ParametreFiche.Dat_code.RecordSource = "select max (para_code) from para"
    frm_ParametreFiche.Dat_code.Refresh
    frm_ParametreFiche.Txt_Code = frm_ParametreFiche.Dat_code.Recordset.Fields(0) + 1
  End If
  
  frm_ParametreFiche.Txt_lib = ""
  frm_ParametreFiche.Txt_nbmois = ""
  
  frm_ParametreFiche.Show vbModal
  
End Sub


Private Sub cmd_Excel_Click()
  
  ExportExcel (dat_para.RecordSource)
  
End Sub

Private Sub cmd_Modifier_Click()

  stb_ParametreListe.Panels(3) = "MODIFICATION"

  Load frm_ParametreFiche
  frm_ParametreFiche.Txt_Code = dat_para.Recordset("PARA_CODE")
  frm_ParametreFiche.Txt_lib = dat_para.Recordset("PARA_LIB")
  frm_ParametreFiche.Txt_nbmois = dat_para.Recordset("PARA_NBMOIS")
  frm_ParametreFiche.Show vbModal

End Sub


Private Sub cmd_Quitter_Click()
  Unload Me
End Sub

Private Sub cmd_Supprimer_Click()
  
  stb_ParametreListe.Panels(3) = "SUPPRESSION"

  Load frm_ParametreFiche
  frm_ParametreFiche.Txt_Code = dat_para.Recordset("PARA_CODE")
  frm_ParametreFiche.Txt_lib = dat_para.Recordset("PARA_LIB")
  frm_ParametreFiche.Txt_nbmois = dat_para.Recordset("PARA_NBMOIS")
  
  frm_ParametreFiche.Txt_Code.Enabled = False
  frm_ParametreFiche.Txt_lib.Enabled = False
  frm_ParametreFiche.Txt_nbmois.Enabled = False
  frm_ParametreFiche.Show vbModal

End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
  
  Dim colonne As String
  Dim classem As String
  
  colonne = DBGrid1.Columns(ColIndex).DataField
  Call Tri(dat_para, DBGrid1, colonne)
      
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(dat_para.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_ParametreListe.Panels(4).Text = "Classement " & classem
  
End Sub


Private Sub Form_Load()

dat_para.DatabaseName = CheminBase
dat_para.RecordSource = "SELECT * FROM PARA"
dat_para.Refresh

End Sub

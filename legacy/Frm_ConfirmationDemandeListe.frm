VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_ConfirmationDemandeListe 
   Caption         =   "Liste des confirmations de la demande"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   ControlBox      =   0   'False
   Icon            =   "Frm_ConfirmationDemandeListe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   9285
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Dat_MaxAnnee 
      Caption         =   "Dat_MaxAnnee"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Dat_reno 
      Caption         =   "Dat_reno"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSDBGrid.DBGrid DBG_ConfDemListe 
      Bindings        =   "Frm_ConfirmationDemandeListe.frx":08CA
      Height          =   2535
      Left            =   120
      OleObjectBlob   =   "Frm_ConfirmationDemandeListe.frx":08E1
      TabIndex        =   13
      Top             =   1920
      Width           =   9015
   End
   Begin VB.Frame fra_Dossiers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.Data Dat_demandes 
         Caption         =   "Dat_demandes"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   4680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data Dat_dossiers 
         Caption         =   "Dat_dossiers"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8280
         Picture         =   "Frm_ConfirmationDemandeListe.frx":1BB1
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   170
         Width           =   375
      End
      Begin VB.CommandButton cmd_SupprimerConf 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         Picture         =   "Frm_ConfirmationDemandeListe.frx":1CB3
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprimer la confirmation de demande sélectionnée"
         Top             =   170
         Width           =   375
      End
      Begin VB.CommandButton cmd_ModifierConf 
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         Picture         =   "Frm_ConfirmationDemandeListe.frx":1DB5
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modifier les informations de la confirmation de demande sélectionnée"
         Top             =   170
         Width           =   375
      End
      Begin VB.CommandButton Cmd_CreerConf 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         Picture         =   "Frm_ConfirmationDemandeListe.frx":1EB7
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Créer une confirmation de demande"
         Top             =   170
         Width           =   375
      End
   End
   Begin VB.Frame Fra_Identification 
      Caption         =   "Identification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   9015
      Begin VB.Data Dat_demandeur 
         Caption         =   "Dat_demandeur"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   6600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox Cmb_NumDossier 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox Cmb_NumDemande 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Txt_NomDemandeur 
         Height          =   285
         Left            =   3480
         TabIndex        =   8
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Txt_PrenomDemandeur 
         Height          =   285
         Left            =   5880
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTP_DateDemande 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   6
         ToolTipText     =   "Date de la demande"
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyy"
         Format          =   24576003
         CurrentDate     =   36287
      End
      Begin VB.Label lbl_NumDossier 
         Caption         =   "Numéro de dossier"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Lbl_NumDemande 
         Caption         =   "Numéro de demande"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Lbl_DateDemande 
         Caption         =   "Date de la demande"
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
   End
   Begin MSComctlLib.StatusBar stb_ConfDemListe 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   4605
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "17/11/99"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11:31"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Frm_ConfirmationDemandeListe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmb_NumDemande_Click()

Dim sql As String

  If Cmb_NumDemande.Text = "" Then
    If Dat_demandes.Recordset.RecordCount = 0 Then
        MsgBox "Aucune demande n'est disponible pour ce dossier.", vbOKOnly, "ERREUR !"
        Cmd_CreerConf.Enabled = False
        cmd_ModifierConf.Enabled = False
        cmd_SupprimerConf.Enabled = False
        Cmb_NumDossier.SetFocus
    Else
        MsgBox "Vous devez choisir un numéro de demande", vbOKOnly, "Erreur de saisie"
        Cmd_CreerConf.Enabled = False
        cmd_ModifierConf.Enabled = False
        cmd_SupprimerConf.Enabled = False
        Cmb_NumDemande.SetFocus
    End If
  Else
    Cmd_CreerConf.Enabled = True
    cmd_ModifierConf.Enabled = True
    cmd_SupprimerConf.Enabled = True
    'DTP_DateDemande = Dat_demandes.recordset("DEMA_DATE")
    Dat_demandeur.RecordSource = "SELECT INDI_NOM, INDI_PREN FROM INDI WHERE ((DOSS_NUM = " + _
        Iifch(Cmb_NumDossier.Text) + ") AND (INDI_CDFR = " + """DE""" + "))"
    Dat_demandeur.Refresh
    If Dat_demandeur.Recordset.RecordCount <> 0 Then
        If Dat_demandeur.Recordset("INDI_NOM") <> "" Then
            Txt_NomDemandeur = Dat_demandeur.Recordset("INDI_NOM")
        End If
        If Dat_demandeur.Recordset("INDI_PREN") <> "" Then
            Txt_PrenomDemandeur = Dat_demandeur.Recordset("INDI_PREN")
        End If
    End If
    
    sql = "select * from RENO where dema_num=" & Iifent(Cmb_NumDemande)
    Dat_reno.RecordSource = sql
    Dat_reno.Refresh
  End If
  
End Sub


Private Sub Cmb_NumDossier_Click()

  If Cmb_NumDossier.Text = "" Then
      If Dat_Dossiers.Recordset.RecordCount = 0 Then
          MsgBox "Il n'y a aucun dossier de saisi : aucun renouvellement n'est donc possible", vbOKOnly, "ERREUR !"
          cmd_Quitter.SetFocus
      Else
          MsgBox "Vous devez choisir un numéro de dossier", vbOKOnly, "Erreur de saisie"
          Cmb_NumDossier.SetFocus
      End If
      Cmd_CreerConf.Enabled = False
      cmd_ModifierConf.Enabled = False
      cmd_SupprimerConf.Enabled = False
  Else
      Cmd_CreerConf.Enabled = False
      cmd_ModifierConf.Enabled = False
      cmd_SupprimerConf.Enabled = False
      Call ListeDemandesPourCeDossier(Dat_demandes, Cmb_NumDemande, Cmb_NumDossier.Text)
      If Cmb_NumDemande.Enabled = True Then
          Cmb_NumDemande.SetFocus
      End If
  End If

End Sub


Private Sub Cmd_CreerConf_Click()
  
Dim sql As String

  If Dat_reno.Recordset.RecordCount = 5 Then
    MsgBox "Attention :" + Chr(13) + _
           "La demande a déjà été renouvelée 5 fois !", vbExclamation
  End If
  
  stb_ConfDemListe.Panels(4).Text = "CREATION"
  Load Frm_ConfirmationDemandeFiche
  sql = "select count(conf_annee) from reno where dema_num=" & Cmb_NumDemande
  Dat_MaxAnnee.RecordSource = sql
  Dat_MaxAnnee.Refresh
  Frm_ConfirmationDemandeFiche.Txt_annee = Dat_MaxAnnee.Recordset(0) + 1
  Frm_ConfirmationDemandeFiche.Show vbModal
  stb_ConfDemListe.Panels(4).Text = ""
    
End Sub


Private Sub cmd_ModifierConf_Click()
  
  If Not Dat_reno.Recordset.EOF Then
    stb_ConfDemListe.Panels(4).Text = "MODIFICATION"
    Load Frm_ConfirmationDemandeFiche
    Frm_ConfirmationDemandeFiche.Txt_annee = Dat_reno.Recordset("CONF_ANNEE")
    If Not IsNull(Dat_reno.Recordset("RENO_DEMREA")) Then
      Frm_ConfirmationDemandeFiche.Chk_DemReac.Value = Iifent(Dat_reno.Recordset("RENO_DEMREA"))
    End If
    If Not IsNull(Dat_reno.Recordset("RENO_RECMAIT")) Then
      Frm_ConfirmationDemandeFiche.Chk_RecMaintien.Value = Iifent(Dat_reno.Recordset("RENO_RECMAIT"))
    End If
    
    If Dat_reno.Recordset("RENO_DDEMMAAS") <> "" Then
        Frm_ConfirmationDemandeFiche.DTP_DemAs = Dat_reno.Recordset("RENO_DDEMMAAS")
    Else
        Frm_ConfirmationDemandeFiche.DTP_DemAs = Date
        Frm_ConfirmationDemandeFiche.DTP_DemAs = ""
    End If
    If Dat_reno.Recordset("RENO_DRECMAS") <> "" Then
        Frm_ConfirmationDemandeFiche.DTP_RecAs = Dat_reno.Recordset("RENO_DRECMAS")
    Else
        Frm_ConfirmationDemandeFiche.DTP_RecAs = Date
        Frm_ConfirmationDemandeFiche.DTP_RecAs = ""
    End If
    If Dat_reno.Recordset("RENO_DDEMMPS") <> "" Then
        Frm_ConfirmationDemandeFiche.DTP_DemPsy = Dat_reno.Recordset("RENO_DDEMMPS")
    Else
        Frm_ConfirmationDemandeFiche.DTP_DemPsy = Date
        Frm_ConfirmationDemandeFiche.DTP_DemPsy = ""
    End If
    If Dat_reno.Recordset("RENO_DRECMPS") <> "" Then
        Frm_ConfirmationDemandeFiche.DTP_RecPsy = Dat_reno.Recordset("RENO_DRECMPS")
    Else
        Frm_ConfirmationDemandeFiche.DTP_RecPsy = Date
        Frm_ConfirmationDemandeFiche.DTP_RecPsy = ""
    End If
    
    Frm_ConfirmationDemandeFiche.Show vbModal
    stb_ConfDemListe.Panels(4).Text = ""
  End If
  
End Sub


Private Sub cmd_Quitter_Click()
  Unload Me
End Sub


Private Sub cmd_SupprimerConf_Click()

Dim sql As String
  
  If Not Dat_reno.Recordset.EOF Then
    sql = "select count(conf_annee) from reno where dema_num=" & Cmb_NumDemande
    Dat_MaxAnnee.RecordSource = sql
    Dat_MaxAnnee.Refresh
    If Dat_MaxAnnee.Recordset(0) > Dat_reno.Recordset("conf_annee") Then
      MsgBox "Vous ne pouvez pas supprimer cette confirmation de demande." + Chr(13) + _
             "Une confirmation postérieure existe !", vbExclamation
      Exit Sub
    Else
      stb_ConfDemListe.Panels(4).Text = "SUPPRESSION"
      Load Frm_ConfirmationDemandeFiche
      Frm_ConfirmationDemandeFiche.Txt_annee.Enabled = False
      Frm_ConfirmationDemandeFiche.Chk_DemReac.Enabled = False
      Frm_ConfirmationDemandeFiche.Chk_RecMaintien.Enabled = False
      Frm_ConfirmationDemandeFiche.DTP_DemAs.Enabled = False
      Frm_ConfirmationDemandeFiche.DTP_DemPsy.Enabled = False
      Frm_ConfirmationDemandeFiche.DTP_RecAs.Enabled = False
      Frm_ConfirmationDemandeFiche.DTP_RecPsy.Enabled = False
      
      Frm_ConfirmationDemandeFiche.Txt_annee = Dat_reno.Recordset("CONF_ANNEE")
      Frm_ConfirmationDemandeFiche.Chk_DemReac.Value = Iifent(Dat_reno.Recordset("RENO_DEMREA"))
      Frm_ConfirmationDemandeFiche.Chk_RecMaintien.Value = Iifent(Dat_reno.Recordset("RENO_RECMAIT"))
      
      If Dat_reno.Recordset("RENO_DDEMMAAS") <> "" Then
          Frm_ConfirmationDemandeFiche.DTP_DemAs = Dat_reno.Recordset("RENO_DDEMMAAS")
      Else
          Frm_ConfirmationDemandeFiche.DTP_DemAs = Date
          Frm_ConfirmationDemandeFiche.DTP_DemAs = ""
      End If
      If Dat_reno.Recordset("RENO_DRECMAS") <> "" Then
          Frm_ConfirmationDemandeFiche.DTP_RecAs = Dat_reno.Recordset("RENO_DRECMAS")
      Else
          Frm_ConfirmationDemandeFiche.DTP_RecAs = Date
          Frm_ConfirmationDemandeFiche.DTP_RecAs = ""
      End If
      If Dat_reno.Recordset("RENO_DDEMMPS") <> "" Then
          Frm_ConfirmationDemandeFiche.DTP_DemPsy = Dat_reno.Recordset("RENO_DDEMMPS")
      Else
          Frm_ConfirmationDemandeFiche.DTP_DemPsy = Date
          Frm_ConfirmationDemandeFiche.DTP_DemPsy = ""
      End If
      If Dat_reno.Recordset("RENO_DRECMPS") <> "" Then
          Frm_ConfirmationDemandeFiche.DTP_RecPsy = Dat_reno.Recordset("RENO_DRECMPS")
      Else
          Frm_ConfirmationDemandeFiche.DTP_RecPsy = Date
          Frm_ConfirmationDemandeFiche.DTP_RecPsy = ""
      End If
      
      Frm_ConfirmationDemandeFiche.Show vbModal
      stb_ConfDemListe.Panels(4).Text = ""
    End If
  End If
End Sub

Private Sub DBG_ConfDemListe_HeadClick(ByVal ColIndex As Integer)

  Dim colonne As String
  Dim classem As String
  
  colonne = DBG_ConfDemListe.Columns(ColIndex).DataField
  Call Tri(Dat_reno, DBG_ConfDemListe, colonne)
      
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_reno.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_ConfDemListe.Panels(4).Text = "Classement " & classem
  
End Sub


Private Sub Form_Load()

  Dat_Dossiers.DatabaseName = CheminBase
  Dat_demandes.DatabaseName = CheminBase
  Dat_demandeur.DatabaseName = CheminBase
  Dat_reno.DatabaseName = CheminBase
  Dat_MaxAnnee.DatabaseName = CheminBase
  
  Call ListeDossiers(Dat_Dossiers, Cmb_NumDossier)
  Call ListeDemandesPourCeDossier(Dat_demandes, Cmb_NumDemande, Cmb_NumDossier.Text)
  
End Sub



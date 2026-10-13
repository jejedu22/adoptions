VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_PlacementGestion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion des placements"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_PlacementGestion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fra_Enfants 
      Caption         =   "Enfants"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   39
      Top             =   4440
      Width           =   9255
      Begin VB.CommandButton cmdAdoptions 
         Height          =   375
         Index           =   0
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Saisir la date d'envoi du dossier au Tribunal de Grande Instance"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   2655
      End
      Begin VB.CommandButton cmd_ModifierEnfant 
         Height          =   375
         Left            =   720
         Picture         =   "frm_PlacementGestion.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Modifier l'enfant sélectionné"
         Top             =   240
         Width           =   375
      End
      Begin VB.Data Dat_Placements 
         Caption         =   "Dat_Placements"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1200
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data Dat_TGI 
         Caption         =   "Dat_TGI"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Data Dat_Oeuvres 
         Caption         =   "Dat_Oeuvres"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1440
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton Cmd_SurveillancePlacement 
         Caption         =   "Surveillances du placement ..."
         Height          =   375
         Left            =   6120
         TabIndex        =   15
         ToolTipText     =   "Fenêtre d'exécution d'une surveillance de placement"
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmd_SupprimerEnfant 
         Height          =   375
         Left            =   1320
         Picture         =   "frm_PlacementGestion.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Supprimer l'enfant sélectionné"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_CreerEnfant 
         Height          =   375
         Left            =   120
         Picture         =   "frm_PlacementGestion.frx":0ACE
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ajouter un nouvel enfant"
         Top             =   240
         Width           =   375
      End
      Begin VB.Data Dat_EnfantsPlaces 
         Caption         =   "Dat_EnfantsPlacés"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ENPL"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSDBGrid.DBGrid grd_EnfantsPlaces 
         Bindings        =   "frm_PlacementGestion.frx":0BD0
         Height          =   1215
         Left            =   120
         OleObjectBlob   =   "frm_PlacementGestion.frx":0BF0
         TabIndex        =   12
         ToolTipText     =   "Liste des enfants placés"
         Top             =   720
         Width           =   9015
      End
   End
   Begin VB.Frame Fra_Jugement 
      Caption         =   "Jugement"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      TabIndex        =   34
      Top             =   3050
      Width           =   8055
      Begin VB.ComboBox Cmb_TGIAccueil 
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Libellé et code du T.G.I."
         Top             =   960
         Width           =   2655
      End
      Begin VB.ComboBox Cmb_TGIOrigine 
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Libellé et code du T.G.I."
         Top             =   480
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DTP_DateJugementAccueil 
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         ToolTipText     =   "Date du jugement dans le pays d'accueil"
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58720259
         CurrentDate     =   36300
      End
      Begin MSComCtl2.DTPicker DTP_DateJugementOrigine 
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         ToolTipText     =   "Date du jugement dans le pays d'origine"
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58720259
         CurrentDate     =   36300
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FF0000&
         X1              =   600
         X2              =   7200
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF0000&
         X1              =   4200
         X2              =   4200
         Y1              =   120
         Y2              =   1320
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         X1              =   2160
         X2              =   2160
         Y1              =   120
         Y2              =   1320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         X1              =   600
         X2              =   7200
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         X1              =   600
         X2              =   7200
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Lbl_NomCodeTGI 
         Caption         =   "Nom et Code T.G.I."
         Height          =   255
         Left            =   4560
         TabIndex        =   38
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Lbl_DateJugement 
         Caption         =   "Date du jugement"
         Height          =   255
         Left            =   2400
         TabIndex        =   37
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Lbl_PaysAccueil 
         Caption         =   "Pays d'accueil"
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Lbl_PaysOrigine 
         Caption         =   "Pays d'origine"
         Height          =   255
         Left            =   600
         TabIndex        =   35
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   26
      Top             =   0
      Width           =   9255
      Begin VB.Data Dat_Natures 
         Caption         =   "Dat_Natures"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   5760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   195
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   8640
         Picture         =   "frm_PlacementGestion.frx":1E31
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   8040
         Picture         =   "frm_PlacementGestion.frx":1F33
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
      Begin VB.Data Dat_DemandesDossier 
         Caption         =   "Dat_DemandesDossier"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   165
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data Dat_Demandeur 
         Caption         =   "Dat_Demandeur"
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
         Top             =   165
         Visible         =   0   'False
         Width           =   2535
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
      TabIndex        =   19
      Top             =   600
      Width           =   9255
      Begin VB.Data Dat_demandes 
         Caption         =   "Dat_demandes"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   7200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Data Dat_Dossiers 
         Caption         =   "Dat_Dossiers"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   315
         Left            =   3120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   0
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.TextBox Txt_NumeroPlacement 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         TabIndex        =   40
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Txt_NomDemandeur 
         Height          =   285
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   21
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Txt_PrenomDemandeur 
         Height          =   285
         Left            =   5880
         MaxLength       =   20
         TabIndex        =   20
         Top             =   240
         Width           =   2415
      End
      Begin VB.ComboBox Cmb_NumDossier 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Choisissez un numéro de dossier"
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox Cmb_NumDemande 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Choisissez un numéro de demande"
         Top             =   600
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTP_DateDemande 
         Height          =   255
         Left            =   6960
         TabIndex        =   22
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58720259
         CurrentDate     =   36287
      End
      Begin VB.Label Lbl_NumeroPlacement 
         Caption         =   "Placement n°"
         Height          =   255
         Left            =   3480
         TabIndex        =   41
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl_NumDossier 
         Caption         =   "Numéro de dossier"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Lbl_NumDemande 
         Caption         =   "Numéro de demande"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Lbl_DateDemande 
         Caption         =   "Date de la demande"
         Height          =   255
         Left            =   5520
         TabIndex        =   23
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Fra_Placement 
      Caption         =   "Placement"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   9255
      Begin MSComCtl2.DTPicker DTP_DateConseilFamille 
         Height          =   255
         Left            =   6840
         TabIndex        =   7
         ToolTipText     =   "Date à laquelle le conseil de famille s'est réuni"
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58720259
         CurrentDate     =   36300
      End
      Begin VB.ComboBox Cmb_OrigineEthnique 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Origine ethnique des enfants placés"
         Top             =   1080
         Width           =   3255
      End
      Begin VB.ComboBox Cmb_Oeuvre 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Oeuvre ayant permi le placement"
         Top             =   720
         Width           =   2655
      End
      Begin VB.ComboBox Cmb_PupilleEtat 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Indique si les enfants placés sont pupilles de l'Etat ou non"
         Top             =   720
         Width           =   495
      End
      Begin VB.ComboBox Cmb_NaturePlacement 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Nature du placement"
         Top             =   240
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DTP_DatePlacement 
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         ToolTipText     =   "Date à laquelle le placement est effectué"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58720259
         CurrentDate     =   36300
      End
      Begin VB.Label Lbl_DateConseilFamille 
         Caption         =   "Date du conseil de famille"
         Height          =   255
         Left            =   4920
         TabIndex        =   33
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Lbl_OrigineEthnique 
         Caption         =   "Origine Ethnique"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Lbl_Oeuvre 
         Caption         =   "Oeuvre"
         Height          =   255
         Left            =   4920
         TabIndex        =   31
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Lbl_PupilleEtat 
         Caption         =   "Pupille de l'Etat"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Lbl_NaturePlacement 
         Caption         =   "Nature du placement"
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Lbl_DatePlacement 
         Caption         =   "Date du placement"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.StatusBar stb_Placement 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   6570
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "02/10/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "09:58"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10548
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_PlacementGestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum IndexCommandButton 'Index des contrôles CommandButton
    cmdENVOI_TGI = 0
End Enum

Private mbooValidationRefusée As Boolean

Private QuitterCroix As Boolean
Private AnnulerLaCreation As Boolean
Private ValidéParMAJEnfant As Boolean
Private ExisteDéjàUnPlacement As Boolean


Private Sub Cmb_NaturePlacement_Gotfocus()
    stb_Placement.SimpleText = "Choisissez la nature du placement"
End Sub

Private Sub Cmb_NumDemande_Click()
Const pkstrNomProcédure As String = "Cmb_NumDemande_Click"
On Error GoTo Err_Cmb_NumDemande_Click

    If Cmb_NumDemande.Text = "" Then
        If Dat_DemandesDossier.Recordset.RecordCount = 0 Then
            MsgBox "Aucune demande n'est disponible pour ce dossier.", vbOKOnly, "ERREUR !"
            Fra_Placement.Enabled = False
            Fra_Jugement.Enabled = False
            fra_enfants.Enabled = False
            Cmb_NumDossier.SetFocus
        Else
            MsgBox "Vous devez choisir un numéro de demande", vbOKOnly, "Erreur de saisie"
            Cmb_NumDemande.SetFocus
        End If
    Else
        Fra_Placement.Enabled = True
        Fra_Jugement.Enabled = True
        fra_enfants.Enabled = True
        
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
        Dat_DemandesDossier.RecordSource = "SELECT DEMA_DATE FROM DEMA WHERE DEMA_NUM=" + Cmb_NumDemande.Text
        Dat_DemandesDossier.Refresh
        DTP_DateDemande = Dat_DemandesDossier.Recordset("DEMA_DATE")
        Dat_Placements.RecordSource = "SELECT COUNT (*) AS Nombre FROM PLAC WHERE DEMA_NUM=" + Cmb_NumDemande.Text
        Dat_Placements.Refresh
        If Dat_Placements.Recordset.Fields(0) <> 0 Then
        
        Select Case stb_Placement.Panels(4).Text
            Case "MODIFICATION"
                ExisteDéjàUnPlacement = True
                Dat_Placements.RecordSource = "SELECT * FROM PLAC WHERE DEMA_NUM = " + Cmb_NumDemande.Text
                Dat_Placements.Refresh
                Txt_NumeroPlacement = Dat_Placements.Recordset("PLAC_NUM")
                DTP_DatePlacement = Dat_Placements.Recordset("PLAC_DATE")
                If Dat_Placements.Recordset("PLAC_DCONS") <> "" Then
                    DTP_DateConseilFamille = Dat_Placements.Recordset("PLAC_DCONS")
                Else
                    DTP_DateConseilFamille = Date
                    DTP_DateConseilFamille = ""
                End If
                Cmb_NaturePlacement.ListIndex = PresentDansListBox(Cmb_NaturePlacement, "<" + CStr(Dat_Placements.Recordset("NATU_CODE")) + ">")
                If IsNull(Dat_Placements.Recordset("oret_code")) Then
                  Cmb_OrigineEthnique.ListIndex = 0
                Else
                  Cmb_OrigineEthnique.ListIndex = PresentDansListBox(Cmb_OrigineEthnique, "<" + CStr(Dat_Placements.Recordset("ORET_CODE")) + ">")
                End If
                If IsNull(Dat_Placements.Recordset("PLAC_PUPILL")) Then
                  Cmb_PupilleEtat.ListIndex = 0
                Else
                  Cmb_PupilleEtat.ListIndex = PresentDansListBox(Cmb_PupilleEtat, Dat_Placements.Recordset("PLAC_PUPILL"))
                End If
                If IsNull(Dat_Placements.Recordset("OEPL_CODE")) Then
                  Cmb_Oeuvre.ListIndex = 0
                Else
                  Cmb_Oeuvre.ListIndex = PresentDansListBox(Cmb_Oeuvre, "<" + CStr(Dat_Placements.Recordset("OEPL_CODE")) + ">")
                End If
                If Dat_Placements.Recordset("JUOR_DJUGTPO") <> "" Then
                    DTP_DateJugementOrigine = Dat_Placements.Recordset("JUOR_DJUGTPO")
                Else
                    DTP_DateJugementOrigine = Date
                    DTP_DateJugementOrigine = ""
                End If
                If Dat_Placements.Recordset("JUAC_DJUGTPA") <> "" Then
                    DTP_DateJugementAccueil = Dat_Placements.Recordset("JUAC_DJUGTPA")
                Else
                    DTP_DateJugementAccueil = Date
                    DTP_DateJugementAccueil = ""
                End If
                If Dat_Placements.Recordset("JUGE_NUMPA") <> "" Then
                    Cmb_TGIAccueil.ListIndex = PresentDansListBox(Cmb_TGIAccueil, "<" + CStr(Dat_Placements.Recordset("JUGE_NUMPA")) + ">")
                End If
                If Dat_Placements.Recordset("JUGE_NUMPO") <> "" Then
                    Cmb_TGIOrigine.ListIndex = PresentDansListBox(Cmb_TGIOrigine, "<" + CStr(Dat_Placements.Recordset("JUGE_NUMPO")) + ">")
                End If
                Dat_EnfantsPlaces.RecordSource = "SELECT * FROM ENPL WHERE PLAC_NUM = " + Iifent(Txt_NumeroPlacement)
                Dat_EnfantsPlaces.Refresh
            
            Case Else
                If ((stb_Placement.Panels(4).Text <> "VISUALISATION") And (stb_Placement.Panels(4).Text <> "SUPPRESSION")) Then
                    MsgBox "Un placement est déjà réalisé pour cette demande." + Chr(10) + _
                        "Vous ne pouvez que consulter les données relatives à ce placement.", vbOKOnly, "Placement déjà réalisé"
                    cmd_Valider.Enabled = False
                    Cmd_SurveillancePlacement.Enabled = True
                End If
                
                ExisteDéjàUnPlacement = True
                Dat_Placements.RecordSource = "SELECT * FROM PLAC WHERE DEMA_NUM = " + Cmb_NumDemande.Text
                Dat_Placements.Refresh
                Txt_NumeroPlacement = Dat_Placements.Recordset("PLAC_NUM")
                DTP_DatePlacement = Dat_Placements.Recordset("PLAC_DATE")
                
                If Dat_Placements.Recordset("PLAC_DCONS") <> "" Then
                    DTP_DateConseilFamille = Dat_Placements.Recordset("PLAC_DCONS")
                Else
                    DTP_DateConseilFamille = Date
                    DTP_DateConseilFamille = ""
                End If
                
                Cmb_NaturePlacement.ListIndex = PresentDansListBox(Cmb_NaturePlacement, "<" + CStr(Dat_Placements.Recordset("NATU_CODE")) + ">")
                If IsNull(Dat_Placements.Recordset("ORET_CODE")) Then
                  Cmb_OrigineEthnique.ListIndex = 0
                Else
                  Cmb_OrigineEthnique.ListIndex = PresentDansListBox(Cmb_OrigineEthnique, "<" + CStr(Dat_Placements.Recordset("ORET_CODE")) + ">")
                End If
                If IsNull(Dat_Placements.Recordset("PLAC_PUPILL")) Then
                  Cmb_PupilleEtat.ListIndex = 0
                Else
                  Cmb_PupilleEtat.ListIndex = PresentDansListBox(Cmb_PupilleEtat, Dat_Placements.Recordset("PLAC_PUPILL"))
                End If
                If IsNull(Dat_Placements.Recordset("OEPL_CODE")) Then
                  Cmb_Oeuvre.ListIndex = 0
                Else
                  Cmb_Oeuvre.ListIndex = PresentDansListBox(Cmb_Oeuvre, "<" + CStr(Dat_Placements.Recordset("OEPL_CODE")) + ">")
                End If
                
                If Dat_Placements.Recordset("JUOR_DJUGTPO") <> "" Then
                    DTP_DateJugementOrigine = Dat_Placements.Recordset("JUOR_DJUGTPO")
                Else
                    DTP_DateJugementOrigine = Date
                    DTP_DateJugementOrigine = ""
                End If
                
                If Dat_Placements.Recordset("JUAC_DJUGTPA") <> "" Then
                    DTP_DateJugementAccueil = Dat_Placements.Recordset("JUAC_DJUGTPA")
                Else
                    DTP_DateJugementAccueil = Date
                    DTP_DateJugementAccueil = ""
                End If
                
                If Dat_Placements.Recordset("JUGE_NUMPA") <> "" Then
                    Cmb_TGIAccueil.ListIndex = PresentDansListBox(Cmb_TGIAccueil, "<" + CStr(Dat_Placements.Recordset("JUGE_NUMPA")) + ">")
                End If
                If Dat_Placements.Recordset("JUGE_NUMPO") <> "" Then
                    Cmb_TGIOrigine.ListIndex = PresentDansListBox(Cmb_TGIOrigine, "<" + CStr(Dat_Placements.Recordset("JUGE_NUMPO")) + ">")
                End If
                
                Cmb_NaturePlacement.Enabled = False
                Cmb_TGIAccueil.Enabled = False
                Cmb_TGIOrigine.Enabled = False
                Cmb_PupilleEtat.Enabled = False
                Cmb_OrigineEthnique.Enabled = False
                Cmb_Oeuvre.Enabled = False
                DTP_DateConseilFamille.Enabled = False
                DTP_DateJugementAccueil.Enabled = False
                DTP_DateJugementOrigine.Enabled = False
                DTP_DatePlacement.Enabled = False
                cmd_CreerEnfant.Enabled = False
                cmd_ModifierEnfant.Enabled = False
                cmd_SupprimerEnfant.Enabled = False
                Dat_EnfantsPlaces.RecordSource = "SELECT * FROM ENPL WHERE PLAC_NUM = " + Iifent(Txt_NumeroPlacement)
                Dat_EnfantsPlaces.Refresh
        End Select
        
        Else
            cmd_Valider.Enabled = False
            Cmd_SurveillancePlacement.Enabled = False
            ExisteDéjàUnPlacement = False
            Txt_NumeroPlacement = NouveauCode("PLAC", "PLAC_NUM")
            DTP_DatePlacement = Date
            DTP_DateConseilFamille = Date
            DTP_DateConseilFamille = ""
            DTP_DateJugementAccueil = Date
            DTP_DateJugementAccueil = ""
            DTP_DateJugementOrigine = Date
            DTP_DateJugementOrigine = ""
            Cmb_NaturePlacement.ListIndex = 0
            Cmb_Oeuvre.ListIndex = 0
            Cmb_PupilleEtat.ListIndex = 0
            Cmb_OrigineEthnique.ListIndex = 0
            ListeTGI Cmb_TGIAccueil
            ListeTGI Cmb_TGIOrigine
            Cmb_NaturePlacement.Enabled = True
            Cmb_TGIAccueil.Enabled = True
            Cmb_TGIOrigine.Enabled = True
            Cmb_PupilleEtat.Enabled = True
            Cmb_OrigineEthnique.Enabled = True
            Cmb_Oeuvre.Enabled = True
            DTP_DateConseilFamille.Enabled = True
            DTP_DateJugementAccueil.Enabled = True
            DTP_DateJugementOrigine.Enabled = True
            DTP_DatePlacement.Enabled = True
            cmd_CreerEnfant.Enabled = True
            cmd_SupprimerEnfant.Enabled = True
            Dat_EnfantsPlaces.RecordSource = "SELECT * FROM ENPL WHERE PLAC_NUM = " + Iifent(Txt_NumeroPlacement)
            Dat_EnfantsPlaces.Refresh
        End If
    End If
    
    Exit Sub
    
Err_Cmb_NumDemande_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub Cmb_NumDemande_LostFocus()
    If Cmb_NumDemande.Text = "" Then
        If Dat_DemandesDossier.Recordset.RecordCount = 0 Then
            MsgBox "Aucune demande n'est disponible pour ce dossier.", vbOKOnly, "ERREUR !"
            Cmb_NumDossier.SetFocus
        Else
            MsgBox "Vous devez choisir un numéro de demande", vbOKOnly, "Erreur de saisie"
            Cmb_NumDemande.SetFocus
        End If
    Else
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
        Dat_DemandesDossier.RecordSource = "SELECT DEMA_DATE FROM DEMA WHERE DEMA_NUM=" + Cmb_NumDemande.Text
        Dat_DemandesDossier.Refresh
        DTP_DateDemande = Dat_DemandesDossier.Recordset("DEMA_DATE")
    End If
End Sub

Private Sub Cmb_NumDossier_Click()
    If Cmb_NumDossier.Text = "" Then
        If frm_DossierGestion.Dat_dossiers.Recordset.RecordCount = 0 Then
            MsgBox "Il n'y a aucun dossier de saisi : aucune action n'est donc possible", vbOKOnly, "ERREUR !"
            Fra_Placement.Enabled = False
            Fra_Jugement.Enabled = False
            fra_enfants.Enabled = False
            cmd_Quitter.SetFocus
        Else
            MsgBox "Vous devez choisir un numéro de dossier", vbOKOnly, "Erreur de saisie"
            Fra_Placement.Enabled = False
            Fra_Jugement.Enabled = False
            fra_enfants.Enabled = False
            Cmb_NumDossier.SetFocus
        End If
    Else
        Call ListeDemandesPourCeDossier(Dat_demandes, Cmb_NumDemande, Cmb_NumDossier.Text)
        Fra_Placement.Enabled = False
        Fra_Jugement.Enabled = False
        fra_enfants.Enabled = False
        If Cmb_NumDemande.Enabled = True Then
            Cmb_NumDemande.SetFocus
        End If
    End If
End Sub

Private Sub Cmb_NumDossier_LostFocus()
    If Cmb_NumDossier.Text = "" Then
        If frm_DossierGestion.Dat_dossiers.Recordset.RecordCount = 0 Then
            MsgBox "Il n'y a aucun dossier de saisi : aucune action n'est donc possible", vbOKOnly, "ERREUR !"
            cmd_Quitter.SetFocus
        Else
            MsgBox "Vous devez choisir un numéro de dossier", vbOKOnly, "Erreur de saisie"
            Cmb_NumDossier.SetFocus
        End If
    Else
        Call ListeDemandesPourCeDossier(Dat_demandes, Cmb_NumDemande, Cmb_NumDossier.Text)
        Cmb_NumDemande.SetFocus
    End If
End Sub

Private Sub Cmb_Oeuvre_Gotfocus()
    stb_Placement.SimpleText = "Choisissez l'oeuvre de placement"
End Sub

Private Sub Cmb_OrigineEthnique_Gotfocus()
    stb_Placement.SimpleText = "Choisissez l'origine ethnique des enfants de ce placement"
End Sub

Private Sub Cmb_PupilleEtat_Gotfocus()
    stb_Placement.SimpleText = "Choisissez si les enfants de ce placement sont pupilles de l'Etat ou non"
End Sub

Private Sub Cmb_TGIAccueil_Gotfocus()
    stb_Placement.SimpleText = "Choisissez le tribunal pour le pays d'accueil"
End Sub

Private Sub Cmb_TGIOrigine_Gotfocus()
    stb_Placement.SimpleText = "Choisissez le tribunal pour le pays d'origine"
End Sub

Private Sub cmd_CreerEnfant_Click()
Const pkstrNomProcédure As String = "cmd_CreerEnfant_Click"
On Error GoTo Err_cmd_CreerEnfant_Click

'   Si le placement n'existe pas
    If ((Not ValidéParMAJEnfant)) And stb_Placement.Panels(4).Text = "CREATION" Then
'       Proposer sa création
        Select Case modMessages.lngRéponseUtilisateur(mesQUESTION_003)
'           L'utilisateur décide sa création
            Case vbYes
                cmd_Valider_Click
                If mbooValidationRefusée Then Exit Sub
'           L'utilisateur décide l'abandon du traitement
            Case vbNo
                Exit Sub
        End Select
    End If
    'ValidéParMAJEnfant = True
    Load frm_PlacementEnfant
    frm_PlacementEnfant.stb_DemandeEnfant.Panels(4).Text = "CREATION"
    frm_PlacementEnfant.Caption = "Ajout d'un enfant"
    frm_PlacementEnfant.Show vbModal
    
    Dat_EnfantsPlaces.Refresh
    If Dat_EnfantsPlaces.Recordset.RecordCount <> 0 Then
        cmd_Valider.Enabled = True
    Else
        cmd_Valider.Enabled = False
    End If
    
    'Cmd_SurveillancePlacement.Enabled = False
    Exit Sub
    
Err_cmd_CreerEnfant_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_ModifierEnfant_Click()

    If Not Dat_EnfantsPlaces.Recordset.EOF Then
        Load frm_PlacementEnfant
        frm_PlacementEnfant.stb_DemandeEnfant.Panels(4).Text = "MODIFICATION"
        frm_PlacementEnfant.Caption = "Modification d'un enfant"
        frm_PlacementEnfant.Txt_Nom.Enabled = False
        frm_PlacementEnfant.Txt_Prenom.Enabled = False
        frm_PlacementEnfant.DTP_DateNaissance.Enabled = True
        frm_PlacementEnfant.Txt_DepNaissance.Enabled = True
        frm_PlacementEnfant.Txt_LieuNaissance.Enabled = True
        frm_PlacementEnfant.Txt_PaysNaissance.Enabled = True
        frm_PlacementEnfant.Cmb_Sexe.Enabled = True

        frm_PlacementEnfant.Txt_Nom = Dat_EnfantsPlaces.Recordset("ENPL_NOM")
        frm_PlacementEnfant.Txt_Prenom = Dat_EnfantsPlaces.Recordset("ENPL_PREN")
        frm_PlacementEnfant.DTP_DateNaissance = Dat_EnfantsPlaces.Recordset("ENPL_DNAIS")
        If Dat_EnfantsPlaces.Recordset("ENPL_DPNAIS") <> "" Then
            frm_PlacementEnfant.Txt_DepNaissance = Dat_EnfantsPlaces.Recordset("ENPL_DPNAIS")
        Else
            frm_PlacementEnfant.Txt_DepNaissance = ""
        End If
        If Dat_EnfantsPlaces.Recordset("ENPL_LNAISS") <> "" Then
            frm_PlacementEnfant.Txt_LieuNaissance = Dat_EnfantsPlaces.Recordset("ENPL_LNAISS")
        Else
            frm_PlacementEnfant.Txt_LieuNaissance = ""
        End If
        If Dat_EnfantsPlaces.Recordset("ENPL_PAYS") <> "" Then
            frm_PlacementEnfant.Txt_PaysNaissance = Dat_EnfantsPlaces.Recordset("ENPL_PAYS")
        Else
            frm_PlacementEnfant.Txt_PaysNaissance = ""
        End If
        frm_PlacementEnfant.Cmb_Sexe.ListIndex = PresentDansListBox(frm_PlacementEnfant.Cmb_Sexe, Dat_EnfantsPlaces.Recordset("ENPL_SEXE"))
        frm_PlacementEnfant.Show vbModal
        
        Dat_EnfantsPlaces.Refresh
        If Dat_EnfantsPlaces.Recordset.RecordCount <> 0 Then
            cmd_Valider.Enabled = True
        Else
            cmd_Valider.Enabled = False
        End If
    End If
    
End Sub

Private Sub cmd_Quitter_Click()
Const pkstrNomProcédure As String = "cmd_Quitter_Click"
Const pkstrFermetureImpossible As String = "La fermeture de la fenêtre est refusée."
Dim prstPlacement As Recordset
Dim pbooAnnulerTransaction As Boolean
Dim requete As String
On Error GoTo Err_cmd_Quitter_Click

    QuitterCroix = False
    
'   Si le placement n'existait pas à l'ouverture de la fenêtre
    If Not ExisteDéjàUnPlacement Then
    
        Screen.MousePointer = vbHourglass
        
'       Ouvrir une transaction
        modDAO.GérerTransaction daoOUVERTURE
    
'       Supprimer les enfants désignés par le placement, si celui-ci a été créé et s'ils existent
        modDAO.SQLSupprimer daoSUPPRESSION_001, False, pkstrFermetureImpossible, "ENPL", CLng(Txt_NumeroPlacement)
        
'       Si la suppression n'a pas généré d'erreur
        If modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
'           Supprimer le placement s'il a été créé
            modDAO.SQLSupprimer daoSUPPRESSION_001, False, pkstrFermetureImpossible, "PLAC", CLng(Txt_NumeroPlacement)
        End If
        
'       Si la suppression a généré une erreur
        If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
'           L'annuler
            modDAO.GérerTransaction daoANNULATION
            Screen.MousePointer = vbDefault: Exit Sub
        End If
            
'       Valider la transaction
        modDAO.GérerTransaction daoVALIDATION
        
'       Si le placement a été supprimé, en informer l'utilisateur
        If Not modDAO.lngNombreEnregistrementsAffectés() = 0 Then Err.Raise mesMESSAGE_004
        
        Screen.MousePointer = vbDefault
    
    End If
    
    Unload Me
    frm_DossierGestion.Dat_demandes.Refresh
    
    Exit Sub
    
Err_cmd_Quitter_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:     pintTraitement = mesINTERROMPRE
        Case mesMESSAGE_004:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
        Case Else:              pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_SupprimerEnfant_Click()
    If Not Dat_EnfantsPlaces.Recordset.EOF Then
        Load frm_PlacementEnfant
        frm_PlacementEnfant.stb_DemandeEnfant.Panels(4).Text = "SUPPRESSION"
        frm_PlacementEnfant.Caption = "Suppression d'un enfant"
        frm_PlacementEnfant.Txt_Nom.Enabled = False
        frm_PlacementEnfant.Txt_Prenom.Enabled = False
        frm_PlacementEnfant.DTP_DateNaissance.Enabled = False
        frm_PlacementEnfant.Txt_DepNaissance.Enabled = False
        frm_PlacementEnfant.Txt_LieuNaissance.Enabled = False
        frm_PlacementEnfant.Txt_PaysNaissance.Enabled = False
        frm_PlacementEnfant.Cmb_Sexe.Enabled = False

        frm_PlacementEnfant.Txt_Nom = Dat_EnfantsPlaces.Recordset("ENPL_NOM")
        frm_PlacementEnfant.Txt_Prenom = Dat_EnfantsPlaces.Recordset("ENPL_PREN")
        frm_PlacementEnfant.DTP_DateNaissance = Dat_EnfantsPlaces.Recordset("ENPL_DNAIS")
        If Dat_EnfantsPlaces.Recordset("ENPL_DPNAIS") <> "" Then
            frm_PlacementEnfant.Txt_DepNaissance = Dat_EnfantsPlaces.Recordset("ENPL_DPNAIS")
        Else
            frm_PlacementEnfant.Txt_DepNaissance = ""
        End If
        If Dat_EnfantsPlaces.Recordset("ENPL_LNAISS") <> "" Then
            frm_PlacementEnfant.Txt_LieuNaissance = Dat_EnfantsPlaces.Recordset("ENPL_LNAISS")
        Else
            frm_PlacementEnfant.Txt_LieuNaissance = ""
        End If
        If Dat_EnfantsPlaces.Recordset("ENPL_PAYS") <> "" Then
            frm_PlacementEnfant.Txt_PaysNaissance = Dat_EnfantsPlaces.Recordset("ENPL_PAYS")
        Else
            frm_PlacementEnfant.Txt_PaysNaissance = ""
        End If
        frm_PlacementEnfant.Cmb_Sexe.ListIndex = PresentDansListBox(frm_PlacementEnfant.Cmb_Sexe, Dat_EnfantsPlaces.Recordset("ENPL_SEXE"))
        frm_PlacementEnfant.Show vbModal
        
        Dat_EnfantsPlaces.Refresh
        If Dat_EnfantsPlaces.Recordset.RecordCount <> 0 Then
            cmd_Valider.Enabled = True
        Else
            cmd_Valider.Enabled = False
        End If
    End If
End Sub

Private Sub Cmd_SurveillancePlacement_Click()
Const pkstrNomProcédure As String = "Cmd_SurveillancePlacement_Click"
On Error GoTo Err_Cmd_SurveillancePlacement_Click

'   Si le placement n'existe pas
    If (Not ValidéParMAJEnfant) And (cmd_Valider.Enabled = True) And Dat_Placements.Recordset.RecordCount = 0 Then
'       Proposer sa création
        Select Case modMessages.lngRéponseUtilisateur(mesQUESTION_003)
'           L'utilisateur décide sa création
            Case vbYes
                cmd_Valider_Click
                If mbooValidationRefusée Then Exit Sub
'           L'utilisateur décide l'abandon du traitement
            Case vbNo
                Exit Sub
        End Select
        ValidéParMAJEnfant = True
    End If
    
'   Ouvrir la fenêtre permettant le traitement de la surveillance du placement
    With frm_PlacementSurveillance
        .Initialiser CLng(Txt_NumeroPlacement)
        .Show vbModal
    End With
            
    Exit Sub
    
Err_Cmd_SurveillancePlacement_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_Valider_Click()
Const pkstrNomProcédure As String = "cmd_Valider_Click"
Dim prstSurveillances As Recordset
Dim pctlFocus As Control
Dim requete As String
Dim pintItération As Integer
On Error GoTo Err_cmd_Valider_Click
QuitterCroix = False

    mbooValidationRefusée = False
    
    Set pctlFocus = IIf(Cmb_NaturePlacement.Enabled, Cmb_NaturePlacement, cmd_Quitter)
    
'   S'il ne s'agit pas d'une supression, refuser la validation si la nature du placement n'est pas sélectionnée
    If Not stb_Placement.Panels(4).Text = "SUPPRESSION" And Cmb_NaturePlacement.Text = "" Then
        Err.Raise mesSÉLECTION_MANQUANTE
        mbooValidationRefusée = True: Exit Sub
    End If

Select Case stb_Placement.Panels(4).Text
    Case "CREATION"
      Me.MousePointer = vbHourglass
        If Not ValidéParMAJEnfant Then
            requete = "INSERT INTO PLAC(PLAC_NUM,DEMA_NUM,OEPL_CODE,JUGE_NUMPA,JUGE_NUMPO,PLAC_DATE,PLAC_DCONS,NATU_CODE,ORET_CODE,PLAC_PUPILL,JUOR_DJUGTPO,JUAC_DJUGTPA)"
            requete = requete + "VALUES("
            requete = requete + Iifent(Txt_NumeroPlacement) + ","
            requete = requete + Iifent(Cmb_NumDemande.Text) + ","
            requete = requete + Iifent(CodeEntreCotes(Cmb_Oeuvre.Text)) + ","
            requete = requete + Iifent(CodeEntreCotes(Cmb_TGIAccueil.Text)) + ","
            requete = requete + Iifent(CodeEntreCotes(Cmb_TGIOrigine.Text)) + ","
            requete = requete + Iifch(DTP_DatePlacement) + ","
            If DTP_DateConseilFamille <> "" Then
                requete = requete + Iifch(DTP_DateConseilFamille) + ","
            Else
                requete = requete + "Null" + ","
            End If
            requete = requete + Iifent(CodeEntreCotes(Cmb_NaturePlacement.Text)) + ","
            requete = requete + Iifent(CodeEntreCotes(Cmb_OrigineEthnique.Text)) + ","
            requete = requete + Iifch(Cmb_PupilleEtat.Text) + ","
            If DTP_DateJugementOrigine <> "" Then
                requete = requete + Iifch(DTP_DateJugementOrigine) + ","
            Else
                requete = requete + "Null" + ","
            End If
            If DTP_DateJugementAccueil <> "" Then
                requete = requete + Iifch(DTP_DateJugementAccueil) + ");"
            Else
                requete = requete + "Null" + ");"
            End If
            Workspaces(0).Databases(0).Execute requete, dbFailOnError
            ValidéParMAJEnfant = True
            stb_Placement.Panels(4).Text = "MODIFICATION"
            frm_PlacementListe.Dat_Placements.Refresh
        'sinon ne doit jamais se produire
        End If
        
    Case "MODIFICATION"
      Me.MousePointer = vbHourglass
        requete = "UPDATE PLAC SET"
        requete = requete + " DEMA_NUM=" + Iifent(Cmb_NumDemande.Text)
        requete = requete + ",OEPL_CODE=" + Iifent(CodeEntreCotes(Cmb_Oeuvre.Text))
        requete = requete + ",JUGE_NUMPA=" + Iifent(CodeEntreCotes(Cmb_TGIAccueil.Text))
        requete = requete + ",JUGE_NUMPO=" + Iifch(CodeEntreCotes(Cmb_TGIOrigine.Text))
        requete = requete + ",NATU_CODE=" + Iifent(CodeEntreCotes(Cmb_NaturePlacement.Text))
        requete = requete + ",ORET_CODE=" + Iifent(CodeEntreCotes(Cmb_OrigineEthnique.Text))
        requete = requete + ",PLAC_DATE="
        If DTP_DatePlacement <> "" Then
            requete = requete + Iifch(DTP_DatePlacement)
        Else
            requete = requete + "Null"
        End If
        requete = requete + ",PLAC_DCONS="
        If DTP_DateConseilFamille <> "" Then
            requete = requete + Iifch(DTP_DateConseilFamille)
        Else
            requete = requete + "Null"
        End If
        requete = requete + ",PLAC_PUPILL=" + Iifch(Cmb_PupilleEtat.Text)
        requete = requete + ",JUOR_DJUGTPO="
        If DTP_DateJugementOrigine <> "" Then
            requete = requete + Iifch(DTP_DateJugementOrigine)
        Else
            requete = requete + "Null"
        End If
        requete = requete + ",JUAC_DJUGTPA="
        If DTP_DateJugementAccueil <> "" Then
            requete = requete + Iifch(DTP_DateJugementAccueil)
        Else
            requete = requete + "Null"
        End If
        requete = requete + " WHERE PLAC_NUM=" + Iifent(Txt_NumeroPlacement)
        
        Workspaces(0).Databases(0).Execute requete, dbFailOnError
        frm_PlacementListe.Dat_Placements.Refresh
        Unload Me
        
    Case "SUPPRESSION"

'           Ouvrir une transaction
            modDAO.GérerTransaction daoOUVERTURE

'           Lire les surveillances de placement
            Set prstSurveillances = modDAO.rstSQLExtraction(daoSURVEILLANCES, True, gstrVALIDATION_ANNULÉE, dbOpenDynaset, CLng(Txt_NumeroPlacement))
            If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then modDAO.GérerTransaction daoANNULATION: Exit Sub
            
            With prstSurveillances
            
'               Pour chacune
                While Not .EOF
                    
'                   Pour chaque intervenant
                    For pintItération = 1 To 2
'                       Supprimer l'historique
                        modDAO.SQLSupprimer daoDEL_HISTORIQUE_INTERVENANT, True, gstrVALIDATION_ANNULÉE, CLng(Cmb_NumDemande), (pintItération = 2), ![SUPL_DDSUR], ![SUPL_intNumero]
                        If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
                    Next pintItération
                        
'                   Valider l'opération
                    modDAO.ValiderOpérationSurRecordset prstSurveillances, pbooSupprimer:=True, pstrListeMessages:=gstrMESSAGE_SUR_VERROUILLAGE_EN_SUPPRESSION
                    Select Case modDAO.lngDernièreErreurDAO()
'                       L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                        Case daoENREGISTREMENT_VERROUILLÉ
                            modDAO.GérerTransaction daoANNULATION: Exit Sub
'                       L'opération a généré un conflit inattendu
                        Case Is <> daoAUCUNE
                            modDAO.GérerTransaction daoANNULATION
                            Err.Raise mesCONFLIT_ACCÈS
                    End Select

                    .MoveNext
                    
                Wend
                
            End With
            
'           Supprimer les enfants désignés par le placement
            modDAO.SQLSupprimer daoSUPPRESSION_001, True, gstrVALIDATION_ANNULÉE, "ENPL", CLng(Txt_NumeroPlacement)
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
            
'           Supprimer le placement
            modDAO.SQLSupprimer daoSUPPRESSION_001, True, gstrVALIDATION_ANNULÉE, "PLAC", CLng(Txt_NumeroPlacement)
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
            
'           Valider la transaction
            modDAO.GérerTransaction daoVALIDATION
            
        frm_PlacementListe.Dat_Placements.Refresh
        Unload Me
    End Select
  Me.MousePointer = vbDefault
    
  Exit Sub
  
Err_cmd_Valider_Click:
Dim pbooAncienneGestionErreurs As Boolean
Dim pintTraitement As MessagesActions
'Affiche le code d'erreur approprié
    Select Case Err
        Case vbObjectError:             pintTraitement = mesINTERROMPRE
        Case mesSÉLECTION_MANQUANTE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, pctlFocus, "de la nature du placement", "Traitement refusé")
        Case 13:                        pbooAncienneGestionErreurs = True: MsgBox "Valeur incompatible avec le type de colonne", vbOKOnly + vbExclamation
        Case 3022:                      pbooAncienneGestionErreurs = True: MsgBox "Doublon" + ", création impossible", vbOKOnly + vbExclamation
        Case 3200:                      pbooAncienneGestionErreurs = True: MsgBox "Reference" + ", suppression impossible", vbOKOnly + vbExclamation
        Case 3075:                      pbooAncienneGestionErreurs = True: MsgBox "L'un des champs contient une cote, mise à jour impossible", vbOKOnly + vbExclamation
        Case 3315:                      pbooAncienneGestionErreurs = True: MsgBox "Un champ obligatoire n'est pas renseigné, mise à jour impossible", vbOKOnly + vbExclamation
        Case mesCONFLIT_ACCÈS:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, modDAO.lngDernièreErreurDAO())
        Case Else:                      pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
'   S'il s'agit de l'ancienne gestion d'erreurs
    If pbooAncienneGestionErreurs Then
        MousePointer = vbDefault
        Exit Sub
'   Sinon
    Else
        If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    End If

End Sub

Private Sub cmdAdoptions_Click(Index As Integer)
Const pkstrNomProcédure As String = "cmdAdoptions_Click"
Dim prstNombrePlacements As Recordset
On Error GoTo Err_cmdAdoptions_Click
    
    Select Case Index
    
'       Définir la date d'envoi du dossier au TGI
        Case cmdENVOI_TGI
        
'           Vérifier l'existence du placement
            Set prstNombrePlacements = modDAO.rstSQLExtraction(daoNOMBRE, True, "La saisie de la date d'envoi est impossible.", "PLAC", "PLAC_NUM", CLng(Txt_NumeroPlacement))
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then Exit Sub
                
'           Si le placement n'existe pas
            If prstNombrePlacements![Nombre] = 0 Then
'               Proposer sa création
                Select Case modMessages.lngRéponseUtilisateur(mesQUESTION_003)
'                   L'utilisateur décide sa création
                    Case vbYes
                        cmd_Valider_Click
                        If mbooValidationRefusée Then Exit Sub
'                   L'utilisateur décide l'abandon du traitement
                    Case vbNo
                        Exit Sub
                End Select
            End If
            
'           Autoriser la saisie de la date d'envoi du dossier au TGI
            frmEnvoiDossierTGI.Initialiser CLng(Txt_NumeroPlacement)
            frmEnvoiDossierTGI.Show vbModal
            
    End Select
        
    Exit Sub
    
Err_cmdAdoptions_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub DTP_DateConseilFamille_GotFocus()
    stb_Placement.SimpleText = "Entrez la date du conseil de famille"
End Sub

Private Sub DTP_DateJugementAccueil_GotFocus()
    stb_Placement.SimpleText = "Entrez la date du jugement pour le pays d'accueil"
End Sub

Private Sub DTP_DateJugementOrigine_GotFocus()
    stb_Placement.SimpleText = "Entrez la date du jugement pour le pays d'origine"
End Sub

Private Sub DTP_DatePlacement_GotFocus()
    stb_Placement.SimpleText = "Entrez la date du placement"
End Sub

Private Sub Form_Load()
Const pkstrNomProcédure As String = "Form_Load"
On Error GoTo Err_Form_Load

    QuitterCroix = True
    ValidéParMAJEnfant = False
    ExisteDéjàUnPlacement = False
    Dat_dossiers.DatabaseName = CheminBase
    Dat_demandes.DatabaseName = CheminBase
    Dat_EnfantsPlaces.DatabaseName = CheminBase
    Dat_DemandesDossier.DatabaseName = CheminBase
    Dat_demandeur.DatabaseName = CheminBase
    Dat_Oeuvres.DatabaseName = CheminBase
    Dat_TGI.DatabaseName = CheminBase
    Dat_Placements.DatabaseName = CheminBase
    Dat_Natures.DatabaseName = CheminBase
    
    ListeOeuvresPlacement Cmb_Oeuvre
    Cmb_Oeuvre.ListIndex = 0
    ListePupilleEtat Cmb_PupilleEtat
    Cmb_PupilleEtat.ListIndex = 0
    ListeNaturesPlacements Cmb_NaturePlacement
    Cmb_NaturePlacement.ListIndex = 0
    ListeOriginesEthniques Cmb_OrigineEthnique
    Cmb_OrigineEthnique.ListIndex = 0
    ListeTGI Cmb_TGIAccueil
    Cmb_TGIAccueil.ListIndex = 0
    ListeTGI Cmb_TGIOrigine
    Cmb_TGIOrigine.ListIndex = 0
        
    cmd_Valider.Enabled = False
    Dat_EnfantsPlaces.RecordSource = ""
    Dat_EnfantsPlaces.Refresh
        
'   Alimenter le contrôle CommandButton
    cmdAdoptions(cmdENVOI_TGI).Picture = Menu.picImage(adpTEXTE_IMAGE, adpTEXTE_ENVOI_DOSSIER_TGI)
        
    Exit Sub
    
Err_Form_Load:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub





Public Sub ListeOeuvresPlacement(Boite As ComboBox)
'Place la liste des oeuvres de placement dans la combo
    With frm_PlacementGestion.Dat_Oeuvres
        .RecordSource = "SELECT OEPL_CODE, OEPL_DESIGN FROM OEPL ORDER BY OEPL_DESIGN"
        .Refresh
        Boite.Clear
        Boite.AddItem ("")
        While Not .Recordset.EOF
            Boite.AddItem (.Recordset.Fields("OEPL_DESIGN") + " <" + CStr(.Recordset.Fields("OEPL_CODE")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub

Public Sub ListePupilleEtat(Boite As ComboBox)
'Permet la sélection d'une fratrie composée d'enfant étant pupille de l'Etat
'(gràce à la combo)
    Boite.AddItem ("")
    Boite.AddItem ("N")
    Boite.AddItem ("O")
End Sub

Public Sub ListeOriginesEthniques(Boite As ComboBox)
'Place la liste des origines ethniques dans la combo
    With frm_PlacementGestion.Dat_TGI
        .RecordSource = "SELECT ORET_LIBELLE, ORET_CODE FROM ORET ORDER BY ORET_LIBELLE"
        .Refresh
        Boite.Clear
        Boite.AddItem ("")
        While Not .Recordset.EOF
            Boite.AddItem (.Recordset.Fields("ORET_LIBELLE") + " <" + CStr(.Recordset.Fields("ORET_CODE")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub

Public Sub ListeTGI(Boite As ComboBox)
'Place la liste des tribunaux dans la combo
    With frm_PlacementGestion.Dat_TGI
        .RecordSource = "SELECT JUGE_NUM, JUGE_NOM FROM JUGE ORDER BY JUGE_NOM"
        .Refresh
        Boite.Clear
        Boite.AddItem ("")
        While Not .Recordset.EOF
            Boite.AddItem (.Recordset.Fields("JUGE_NOM") + " <" + CStr(.Recordset.Fields("JUGE_NUM")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub

Public Sub ListeNaturesPlacements(Boite As ComboBox)
'Place la liste des natures de placements dans la combo
    With Dat_Natures
        .RecordSource = "SELECT NATU_LIBELLE, NATU_CODE FROM NATU ORDER BY NATU_LIBELLE"
        .Refresh
        Boite.Clear
        Boite.AddItem ("")
        While Not Dat_Natures.Recordset.EOF
            Boite.AddItem (.Recordset.Fields("NATU_LIBELLE") + " <" + CStr(.Recordset.Fields("NATU_CODE")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub

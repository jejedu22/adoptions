VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_PlacementListe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liste des placements déjà réalisés"
   ClientHeight    =   6825
   ClientLeft      =   1440
   ClientTop       =   2220
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_PlacementListe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Dat_Intermediaire 
      Caption         =   "Dat_Intermediaire"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4275
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Cmd_VisualiserPlacement 
      Caption         =   "&Visualiser le placement ..."
      Height          =   375
      Left            =   120
      TabIndex        =   34
      ToolTipText     =   "Permet de visualiser le placement et la surveillance de placement"
      Top             =   6120
      Width           =   2175
   End
   Begin VB.OptionButton Opt_RechercheParEnfant 
      Caption         =   "Option2"
      Height          =   255
      Left            =   600
      TabIndex        =   22
      ToolTipText     =   "Permet une recherche d'un placement lorsque l'on connaît les caractéristiques d'un enfant placé."
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton Opt_RechercheParNum 
      Caption         =   "Option1"
      Height          =   255
      Left            =   600
      TabIndex        =   21
      ToolTipText     =   "Permet une recherche d'un placement sur les critères : [ numéro du dossier / numéro de demande / numéro du placement ]"
      Top             =   720
      Width           =   255
   End
   Begin VB.Data Dat_Placements 
      Caption         =   "Dat_Placements"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4275
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "frm_PlacementListe.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Modifier le placement sélectionné"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "frm_PlacementListe.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Enregistrer un nouveau placement"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "frm_PlacementListe.frx":0ACE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Supprimer le placement sélectionné"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Excel 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   4560
         Picture         =   "frm_PlacementListe.frx":0BD0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Export EXCEL"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8760
         Picture         =   "frm_PlacementListe.frx":0D5A
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
         Picture         =   "frm_PlacementListe.frx":0E5C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Actualiser 
         Default         =   -1  'True
         Height          =   375
         Left            =   7320
         Picture         =   "frm_PlacementListe.frx":0F5E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Rechercher d'après les critères"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_ListePlacementsRealises 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   19
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
            TextSave        =   "28/02/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15:06"
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBGrid.DBGrid grd_PlacementsRealises 
      Bindings        =   "frm_PlacementListe.frx":1060
      Height          =   2655
      Left            =   120
      OleObjectBlob   =   "frm_PlacementListe.frx":107D
      TabIndex        =   17
      ToolTipText     =   "Liste des placements réalisés"
      Top             =   3360
      Width           =   9255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      Height          =   2055
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   9255
      Begin VB.TextBox Txt_NomDemandeur 
         Height          =   285
         Left            =   1560
         TabIndex        =   36
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmd_NomDemandeur 
         Caption         =   "Nom Demandeur"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Txt_Pays 
         Height          =   285
         Left            =   7440
         MaxLength       =   20
         TabIndex        =   16
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Txt_Departement 
         Height          =   285
         Left            =   7440
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTP_DateNaissance 
         Height          =   255
         Left            =   7680
         TabIndex        =   14
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   16384003
         CurrentDate     =   36311
      End
      Begin VB.TextBox Txt_Lieu 
         Height          =   285
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Txt_PrenomEnfant 
         Height          =   285
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   12
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Txt_NomEnfant 
         Height          =   285
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton Cmd_Pays 
         Caption         =   "Pa&ys"
         Height          =   255
         Left            =   6240
         TabIndex        =   31
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Cmd_Lieu 
         Caption         =   "&Lieu"
         Height          =   255
         Left            =   3240
         TabIndex        =   30
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Cmd_Departement 
         Caption         =   "Département"
         Height          =   255
         Left            =   6240
         TabIndex        =   29
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Cmd_DateNaissance 
         Caption         =   "D&ate de naissance"
         Height          =   255
         Left            =   6240
         TabIndex        =   28
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Cmd_Prenom 
         Caption         =   "P&rénom"
         Height          =   255
         Left            =   3240
         TabIndex        =   27
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Cmd_NomEnfant 
         Caption         =   "&Nom"
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Cmd_NumDemande 
         Caption         =   "N° De&mande"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_NumDossier 
         Caption         =   "N° &Dossier"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_NumPlacement 
         Caption         =   "N° &Placement"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Txt_NumDemande 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Txt_NumDossier 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Txt_NumPlacement 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Lbl_RechercheEnfantPlace 
         Caption         =   "Rechercher un enfant placé dans une famille :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   33
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Lbl_RechercheParDossiers 
         Caption         =   "Rechercher dans les dossiers :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   2295
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         X1              =   3000
         X2              =   3000
         Y1              =   240
         Y2              =   1920
      End
   End
   Begin VB.Label Lbl_RechercheParEnfantPlace 
      Caption         =   "Recherche par enfant placé"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      ToolTipText     =   "Permet une recherche d'un placement lorsque l'on connaît les caractéristiques d'un enfant placé."
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Lbl_RechercheParDossDemPlac 
      Caption         =   "Recherche par N° Dossier / N° Demande / N° Placement"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      ToolTipText     =   "Permet une recherche d'un placement sur les critères : [ numéro du dossier / numéro de demande / numéro du placement ]"
      Top             =   720
      Width           =   4215
   End
End
Attribute VB_Name = "frm_PlacementListe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_Actualiser_Click()

Dim sel_numdossier As String
Dim sel_numplacement As String
Dim sel_numdemande As String
Dim sel_nomdemandeur As String
Dim sel_nomenfant As String
Dim sel_prenomenfant As String
Dim sel_datenaissance As String
Dim sel_departement As String
Dim sel_lieu As String
Dim sel_pays As String
    
    If Txt_NumPlacement <> "" Then sel_numplacement = " AND PLAC.PLAC_NUM like """ & Txt_NumPlacement & "*"""
    If Txt_numDossier <> "" Then sel_numdossier = " AND DOSS.DOSS_NUM like """ & Txt_numDossier & "*"""
    If Txt_NumDemande <> "" Then sel_numdemande = " AND DEMA.DEMA_NUM like """ & Txt_NumDemande & "*"""
    If Txt_NomDemandeur <> "" Then sel_nomdemandeur = " AND INDI.INDI_NOM like """ & Txt_NomDemandeur & "*"""
    If Txt_NomEnfant <> "" Then sel_nomenfant = " AND ENPL_NOM like """ & Txt_NomEnfant & "*"""
    If Txt_PrenomEnfant <> "" Then sel_prenomenfant = " AND ENPL_PREN like """ & Txt_PrenomEnfant & "*"""
    If DTP_DateNaissance <> "" Then sel_datenaissance = " AND ENPL_DNAIS like """ & DTP_DateNaissance & "*"""
    If Txt_Departement <> "" Then sel_departement = " AND ENPL_DPNAIS like """ & Txt_Departement & "*"""
    If Txt_Lieu <> "" Then sel_lieu = " AND ENPL_LNAISS like """ & Txt_Lieu & "*"""
    If Txt_Pays <> "" Then sel_pays = " AND ENPL_PAYS like """ & Txt_Pays & "*"""
        
    Dat_Placements.RecordSource = "SELECT DISTINCT PLAC.PLAC_NUM, DEMA.DEMA_NUM, DOSS.DOSS_NUM, " + _
                                  "INDI.INDI_NOM, INDI.INDI_PREN, PLAC.PLAC_DATE, DEMA.DEMA_DATE " + _
                                  "FROM PLAC, DOSS, DEMA, INDI,ENPL " + _
                                  "WHERE DEMA.DEMA_NUM = PLAC.DEMA_NUM AND DEMA.DOSS_NUM = DOSS.DOSS_NUM " + _
                                  "AND ENPL.PLAC_NUM=PLAC.PLAC_NUM " + _
                                  "AND INDI.DOSS_NUM = DOSS.DOSS_NUM AND INDI.INDI_CDFR = ""DE"" " + _
                                  sel_numdossier + _
                                  sel_numdemande + _
                                  sel_numplacement + _
                                  sel_nomdemandeur + _
                                  sel_nomenfant + _
                                  sel_prenomenfant + _
                                  sel_datenaissance + _
                                  sel_departement + _
                                  sel_lieu + _
                                  sel_pays + _
                                  ClauseOrder
    Me.MousePointer = vbHourglass
    Dat_Placements.Refresh
    Me.MousePointer = vbDefault
End Sub

Private Sub cmd_Creer_Click()
    Me.MousePointer = vbHourglass
    stb_ListePlacementsRealises.Panels(4) = "CREATION"  'sc
    Load frm_PlacementGestion
    frm_PlacementGestion.stb_Placement.Panels(4).Text = "CREATION"
    Call ListeDossiers(frm_PlacementGestion.Dat_Dossiers, frm_PlacementGestion.Cmb_NumDossier)
    Call ListeDemandesPourCeDossier(frm_PlacementGestion.Dat_demandes, frm_PlacementGestion.Cmb_NumDemande, frm_PlacementGestion.Cmb_NumDossier.Text)
    frm_PlacementGestion.Txt_NomDemandeur.Enabled = False
    frm_PlacementGestion.Txt_PrenomDemandeur.Enabled = False
    frm_PlacementGestion.DTP_DateDemande.Enabled = False
    frm_PlacementGestion.Caption = "Enregistrement d'un placement"
    Me.MousePointer = vbDefault
    frm_PlacementGestion.Show vbModal
    Cmd_Actualiser_Click
End Sub

Private Sub Cmd_DateNaissance_Click()
    DTP_DateNaissance = Date
    DTP_DateNaissance = ""
    DTP_DateNaissance.SetFocus
End Sub

Private Sub Cmd_Departement_Click()
    Txt_Departement = ""
    Txt_Departement.SetFocus
End Sub

Private Sub cmd_Excel_Click()

  ExportExcel (Dat_Placements.RecordSource)

End Sub

Private Sub cmd_lieu_Click()
    Txt_Lieu = ""
    Txt_Lieu.SetFocus
End Sub

Private Sub cmd_Modifier_Click()

If Not Dat_Placements.Recordset.EOF Then
        Me.MousePointer = vbHourglass
        stb_ListePlacementsRealises.Panels(4) = "MODIFICATION"
        Load frm_PlacementGestion
        frm_PlacementGestion.stb_Placement.Panels(4) = "MODIFICATION"
        frm_PlacementGestion.Cmb_NumDossier.Enabled = False
        frm_PlacementGestion.Cmb_NumDemande.Enabled = False
        frm_PlacementGestion.Txt_NomDemandeur.Enabled = False
        frm_PlacementGestion.Txt_PrenomDemandeur.Enabled = False
        
        Call ListeDossiers(frm_PlacementGestion.Dat_Dossiers, frm_PlacementGestion.Cmb_NumDossier)
        frm_PlacementGestion.Cmb_NumDossier.ListIndex = PresentDansListBox(frm_PlacementGestion.Cmb_NumDossier, Dat_Placements.Recordset("DOSS_NUM"))
        Call ListeDemandesPourCeDossier(frm_PlacementGestion.Dat_demandes, frm_PlacementGestion.Cmb_NumDemande, frm_PlacementGestion.Cmb_NumDossier.Text)
        frm_PlacementGestion.Cmb_NumDemande.ListIndex = PresentDansListBox(frm_PlacementGestion.Cmb_NumDemande, Dat_Placements.Recordset("DEMA_NUM"))
        Dat_Intermediaire.RecordSource = "SELECT INDI_NOM, INDI_PREN FROM INDI WHERE ((DOSS_NUM = " + _
            Iifch(frm_PlacementGestion.Cmb_NumDossier.Text) + ") AND (INDI_CDFR = " + """DE""" + "))"
        Dat_Intermediaire.Refresh
        
        If Dat_Intermediaire.Recordset.RecordCount <> 0 Then
            If Dat_Intermediaire.Recordset("INDI_NOM") <> "" Then
                frm_PlacementGestion.Txt_NomDemandeur = Dat_Intermediaire.Recordset("INDI_NOM")
            End If
            If Dat_Intermediaire.Recordset("INDI_PREN") <> "" Then
                frm_PlacementGestion.Txt_PrenomDemandeur = Dat_Intermediaire.Recordset("INDI_PREN")
            End If
        End If
        
        Dat_Intermediaire.RecordSource = "SELECT DEMA_DATE FROM DEMA WHERE DEMA_NUM = " + frm_PlacementGestion.Cmb_NumDemande.Text
        Dat_Intermediaire.Refresh
        frm_PlacementGestion.DTP_DateDemande = Dat_Intermediaire.Recordset("DEMA_DATE")

        frm_PlacementGestion.Caption = "Modification du placement"
        frm_PlacementGestion.cmd_Valider.Enabled = True
        Me.MousePointer = vbDefault
        frm_PlacementGestion.Show vbModal
        stb_ListePlacementsRealises.Panels(4).Text = ""
    End If

End Sub

Private Sub Cmd_NomDemandeur_Click()
    Txt_NomDemandeur = ""
    Txt_NomDemandeur.SetFocus

End Sub

Private Sub Cmd_NomEnfant_Click()
    Txt_NomEnfant = ""
    Txt_NomEnfant.SetFocus
End Sub

Private Sub Cmd_NumDemande_Click()
    Txt_NumDemande = ""
    Txt_NumDemande.SetFocus
End Sub

Private Sub Cmd_NumDossier_Click()
    Txt_numDossier = ""
    Txt_numDossier.SetFocus
End Sub

Private Sub Cmd_NumPlacement_Click()
    Txt_NumPlacement = ""
    Txt_NumPlacement.SetFocus
End Sub

Private Sub Cmd_Pays_Click()
    Txt_Pays = ""
    Txt_Pays.SetFocus
End Sub

Private Sub Cmd_Prenom_Click()
    Txt_PrenomEnfant = ""
    Txt_PrenomEnfant.SetFocus
End Sub

Private Sub cmd_Quitter_Click()
    Unload Me
End Sub

Private Sub cmd_Supprimer_Click()

    If Not Dat_Placements.Recordset.EOF Then
        Me.MousePointer = vbHourglass
        stb_ListePlacementsRealises.Panels(4) = "SUPPRESSION"
        Load frm_PlacementGestion
        frm_PlacementGestion.stb_Placement.Panels(4) = "SUPPRESSION"
        frm_PlacementGestion.Cmb_NumDossier.Enabled = False
        frm_PlacementGestion.Cmb_NumDemande.Enabled = False
        frm_PlacementGestion.Cmb_PupilleEtat.Enabled = False
        frm_PlacementGestion.Cmb_OrigineEthnique.Enabled = False
        frm_PlacementGestion.Cmb_NaturePlacement.Enabled = False
        frm_PlacementGestion.Cmb_Oeuvre.Enabled = False
        frm_PlacementGestion.Cmb_TGIOrigine.Enabled = False
        frm_PlacementGestion.Cmb_TGIAccueil.Enabled = False
        frm_PlacementGestion.Txt_NomDemandeur.Enabled = False
        frm_PlacementGestion.Txt_PrenomDemandeur.Enabled = False
        frm_PlacementGestion.DTP_DateDemande.Enabled = False
        frm_PlacementGestion.Cmd_SurveillancePlacement.Enabled = False
        frm_PlacementGestion.cmd_CreerEnfant.Enabled = False
        frm_PlacementGestion.cmd_SupprimerEnfant.Enabled = False
        frm_PlacementGestion.cmd_ModifierEnfant.Enabled = False

        Call ListeDossiers(frm_PlacementGestion.Dat_Dossiers, frm_PlacementGestion.Cmb_NumDossier)
        frm_PlacementGestion.Cmb_NumDossier.ListIndex = PresentDansListBox(frm_PlacementGestion.Cmb_NumDossier, Dat_Placements.Recordset("DOSS_NUM"))
        Call ListeDemandesPourCeDossier(frm_PlacementGestion.Dat_demandes, frm_PlacementGestion.Cmb_NumDemande, frm_PlacementGestion.Cmb_NumDossier.Text)
        frm_PlacementGestion.Cmb_NumDemande.ListIndex = PresentDansListBox(frm_PlacementGestion.Cmb_NumDemande, Dat_Placements.Recordset("DEMA_NUM"))
        Dat_Intermediaire.RecordSource = "SELECT INDI_NOM, INDI_PREN FROM INDI WHERE ((DOSS_NUM = " + _
            Iifch(frm_PlacementGestion.Cmb_NumDossier.Text) + ") AND (INDI_CDFR = " + """DE""" + "))"
        Dat_Intermediaire.Refresh
        
        If Dat_Intermediaire.Recordset.RecordCount <> 0 Then
            If Dat_Intermediaire.Recordset("INDI_NOM") <> "" Then
                frm_PlacementGestion.Txt_NomDemandeur = Dat_Intermediaire.Recordset("INDI_NOM")
            End If
            If Dat_Intermediaire.Recordset("INDI_PREN") <> "" Then
                frm_PlacementGestion.Txt_PrenomDemandeur = Dat_Intermediaire.Recordset("INDI_PREN")
            End If
        End If
        
        Dat_Intermediaire.RecordSource = "SELECT DEMA_DATE FROM DEMA WHERE DEMA_NUM = " + frm_PlacementGestion.Cmb_NumDemande.Text
        Dat_Intermediaire.Refresh
        frm_PlacementGestion.DTP_DateDemande = Dat_Intermediaire.Recordset("DEMA_DATE")



        frm_PlacementGestion.Caption = "Suppression du placement"
        frm_PlacementGestion.cmd_Valider.Enabled = True
        Me.MousePointer = vbDefault
        frm_PlacementGestion.Show vbModal
        stb_ListePlacementsRealises.Panels(4).Text = ""
    End If
End Sub

Private Sub cmd_Valider_Click()
    If Not Dat_Placements.Recordset.EOF Then
        Me.Hide
    End If
End Sub

Private Sub Cmd_VisualiserPlacement_Click()
    If Not Dat_Placements.Recordset.EOF Then
        stb_ListePlacementsRealises.Panels(4) = "VISUALISATION"
        Load frm_PlacementGestion
        frm_PlacementGestion.stb_Placement.Panels(4) = "VISUALISATION"
        frm_PlacementGestion.Cmb_NumDossier.Enabled = False
        frm_PlacementGestion.Cmb_NumDemande.Enabled = False
        frm_PlacementGestion.Txt_NomDemandeur.Enabled = False
        frm_PlacementGestion.Txt_PrenomDemandeur.Enabled = False
        frm_PlacementGestion.DTP_DateDemande.Enabled = False
        frm_PlacementGestion.Cmd_SurveillancePlacement.Enabled = True

        Call ListeDossiers(frm_PlacementGestion.Dat_Dossiers, frm_PlacementGestion.Cmb_NumDossier)
        frm_PlacementGestion.Cmb_NumDossier.ListIndex = PresentDansListBox(frm_PlacementGestion.Cmb_NumDossier, Dat_Placements.Recordset("DOSS_NUM"))
        Call ListeDemandesPourCeDossier(frm_PlacementGestion.Dat_demandes, frm_PlacementGestion.Cmb_NumDemande, frm_PlacementGestion.Cmb_NumDossier.Text)
        frm_PlacementGestion.Cmb_NumDemande.ListIndex = PresentDansListBox(frm_PlacementGestion.Cmb_NumDemande, Dat_Placements.Recordset("DEMA_NUM"))
        Dat_Intermediaire.RecordSource = "SELECT INDI_NOM, INDI_PREN FROM INDI WHERE ((DOSS_NUM = " + _
            Iifch(frm_PlacementGestion.Cmb_NumDossier.Text) + ") AND (INDI_CDFR = " + """DE""" + "))"
        Dat_Intermediaire.Refresh
        If Dat_Intermediaire.Recordset.RecordCount <> 0 Then
            If Dat_Intermediaire.Recordset("INDI_NOM") <> "" Then
                frm_PlacementGestion.Txt_NomDemandeur = Dat_Intermediaire.Recordset("INDI_NOM")
            End If
            If Dat_Intermediaire.Recordset("INDI_PREN") <> "" Then
                frm_PlacementGestion.Txt_PrenomDemandeur = Dat_Intermediaire.Recordset("INDI_PREN")
            End If
        End If
        Dat_Intermediaire.RecordSource = "SELECT DEMA_DATE FROM DEMA WHERE DEMA_NUM = " + frm_PlacementGestion.Cmb_NumDemande.Text
        Dat_Intermediaire.Refresh
        frm_PlacementGestion.DTP_DateDemande = Dat_Intermediaire.Recordset("DEMA_DATE")

        frm_PlacementGestion.Caption = "Visualisation du placement"
        frm_PlacementGestion.Show vbModal
        stb_ListePlacementsRealises.Panels(4).Text = ""
    End If
End Sub

Private Sub Form_Load()

    Dat_Placements.DatabaseName = CheminBase
    Dat_Intermediaire.DatabaseName = CheminBase
    Dat_Placements.RecordSource = "SELECT DISTINCT PLAC.PLAC_NUM, DEMA.DEMA_NUM, DOSS.DOSS_NUM, " + _
                                  "INDI.INDI_NOM, INDI.INDI_PREN, PLAC.PLAC_DATE, DEMA.DEMA_DATE " + _
                                  "FROM PLAC, DOSS, DEMA, INDI " + _
                                  "WHERE (DEMA.DEMA_NUM = PLAC.DEMA_NUM AND DEMA.DOSS_NUM = DOSS.DOSS_NUM " + _
                                  "AND INDI.DOSS_NUM = DOSS.DOSS_NUM AND INDI.INDI_CDFR = ""DE"")"
    Dat_Placements.Refresh
    Cmd_NumPlacement.Enabled = False
    Cmd_NumDossier.Enabled = False
    Cmd_NumDemande.Enabled = False
    cmd_NomDemandeur.Enabled = False
    Txt_NumPlacement.Enabled = False
    Txt_numDossier.Enabled = False
    Txt_NumDemande.Enabled = False
    Txt_NomDemandeur.Enabled = False
    
    Cmd_DateNaissance.Enabled = False
    Cmd_Departement.Enabled = False
    Cmd_Lieu.Enabled = False
    Cmd_Pays.Enabled = False
    Cmd_NomEnfant.Enabled = False
    Cmd_Prenom.Enabled = False
    DTP_DateNaissance = Date
    DTP_DateNaissance = ""
    DTP_DateNaissance.Enabled = False
    Txt_Departement.Enabled = False
    Txt_Lieu.Enabled = False
    Txt_Pays.Enabled = False
    Txt_NomEnfant.Enabled = False
    Txt_PrenomEnfant.Enabled = False
End Sub

Private Sub grd_PlacementsRealises_HeadClick(ByVal ColIndex As Integer)

  Dim colonne As String
  Dim classem As String
  
  colonne = grd_PlacementsRealises.Columns(ColIndex).DataField
  Call Tri(Dat_Placements, grd_PlacementsRealises, colonne)
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_Placements.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_ListePlacementsRealises.Panels(4).Text = "Classement " & classem

End Sub

Private Sub Lbl_RechercheParDossDemPlac_Click()
    Opt_RechercheParNum = True
End Sub

Private Sub Lbl_RechercheParEnfantPlace_Click()
    Opt_RechercheParEnfant = True
End Sub

Private Sub Opt_RechercheParEnfant_Click()
    Txt_NumPlacement = ""
    Txt_numDossier = ""
    Txt_NumDemande = ""
    Txt_NomDemandeur = ""
    Cmd_NumPlacement.Enabled = False
    Cmd_NumDossier.Enabled = False
    Cmd_NumDemande.Enabled = False
    cmd_NomDemandeur.Enabled = False
    Txt_NumPlacement.Enabled = False
    Txt_numDossier.Enabled = False
    Txt_NumDemande.Enabled = False
    Txt_NomDemandeur.Enabled = False
    
    Cmd_DateNaissance.Enabled = True
    Cmd_Departement.Enabled = True
    Cmd_Lieu.Enabled = True
    Cmd_Pays.Enabled = True
    Cmd_NomEnfant.Enabled = True
    Cmd_Prenom.Enabled = True
    DTP_DateNaissance.Enabled = True
    Txt_Departement.Enabled = True
    Txt_Lieu.Enabled = True
    Txt_Pays.Enabled = True
    Txt_NomEnfant.Enabled = True
    Txt_PrenomEnfant.Enabled = True
    
    Lbl_RechercheEnfantPlace.ForeColor = &HFF0000
    Lbl_RechercheParDossiers.ForeColor = &H0&
    Cmd_Actualiser_Click
End Sub

Private Sub Opt_RechercheParNum_Click()
    Cmd_NumPlacement.Enabled = True
    Cmd_NumDossier.Enabled = True
    Cmd_NumDemande.Enabled = True
    cmd_NomDemandeur.Enabled = True
    Txt_NumPlacement.Enabled = True
    Txt_numDossier.Enabled = True
    Txt_NumDemande.Enabled = True
    Txt_NomDemandeur.Enabled = True
    
    DTP_DateNaissance = Date
    DTP_DateNaissance = ""
    Txt_Departement = ""
    Txt_Lieu = ""
    Txt_Pays = ""
    Txt_NomEnfant = ""
    Txt_PrenomEnfant = ""
        
    Cmd_DateNaissance.Enabled = False
    Cmd_Departement.Enabled = False
    Cmd_Lieu.Enabled = False
    Cmd_Pays.Enabled = False
    Cmd_NomEnfant.Enabled = False
    Cmd_Prenom.Enabled = False
    DTP_DateNaissance.Enabled = False
    Txt_Departement.Enabled = False
    Txt_Lieu.Enabled = False
    Txt_Pays.Enabled = False
    Txt_NomEnfant.Enabled = False
    Txt_PrenomEnfant.Enabled = False
    
    Lbl_RechercheParDossiers.ForeColor = &HFF0000
    Lbl_RechercheEnfantPlace.ForeColor = &H0&
    Cmd_Actualiser_Click
End Sub

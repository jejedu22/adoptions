VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Canton 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liste des cantons"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   1035
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_Canton.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_Communes 
      Caption         =   "Associer des communes au canton sélectionné"
      Height          =   735
      Left            =   7200
      TabIndex        =   15
      Top             =   3480
      Width           =   1935
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
         Picture         =   "frm_Canton.frx":0442
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
         Picture         =   "frm_Canton.frx":0544
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
         Picture         =   "frm_Canton.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Excel 
         Height          =   375
         Left            =   4560
         Picture         =   "frm_Canton.frx":0748
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Export EXCEL"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "frm_Canton.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprimer le canton sélectionné"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "frm_Canton.frx":09D4
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Créer un nouveau canton"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "frm_Canton.frx":0AD6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modifier les données du canton sélectionné"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   9255
      Begin VB.TextBox Txt_Libelle 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   810
         Width           =   4215
      End
      Begin VB.CommandButton cmd_Libelle 
         Caption         =   "&Libellé"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Num_canton 
         Caption         =   "&N° de canton"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   420
         Width           =   1575
      End
      Begin VB.TextBox Txt_num_canton 
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   420
         Width           =   1095
      End
   End
   Begin VB.Data Dat_Canton 
      Caption         =   "Dat_Canton"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM CANTON ORDER BY code_depart, code_canton"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSDBGrid.DBGrid grd_canton 
      Bindings        =   "frm_Canton.frx":0BD8
      Height          =   4215
      Left            =   1200
      OleObjectBlob   =   "frm_Canton.frx":0BF1
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2160
      Width           =   5655
   End
   Begin MSComctlLib.StatusBar stb_ListeCanton 
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
Attribute VB_Name = "frm_Canton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_Actualiser_Click()

  Dim sel_code_canton As String
  Dim sel_nom_canton As String
  
  Me.MousePointer = vbHourglass
  If Txt_num_canton <> "" Then sel_code_canton = " AND CANT_NUM LIKE """ & Txt_num_canton & "*"""
  If Txt_Libelle <> "" Then sel_nom_canton = " AND CANT_NOM LIKE """ & Txt_Libelle & "*"""
  
  Dat_Canton.RecordSource = "SELECT * FROM CANT WHERE TRUE" + _
                             sel_code_canton + _
                             sel_nom_canton + _
                             " ORDER BY CANT_NUM"
                             
  Dat_Canton.Refresh
  Me.MousePointer = vbDefault
End Sub






Private Sub Cmd_Communes_Click()

  Dim reponse As String

  Dim vntTemp As Variant
  'Dim Code_Dep As String
  Dim requete1 As String
  Dim requete2 As String
  
  If Not Dat_Canton.Recordset.EOF Then
  
    Load frm_CommunesCanton
          
    ' On rempli les textboxs avec les valeurs du canton sélectionnée
    frm_CommunesCanton.Txt_Code = Dat_Canton.Recordset("CANT_NUM")
    frm_CommunesCanton.Txt_Nom = Dat_Canton.Recordset("CANT_NOM")
      
    ' data contenant les villes
    frm_CommunesCanton.Data_Recherche.RecordSource = "SELECT * FROM COMM"
    frm_CommunesCanton.Data_Recherche.Refresh
      
    ' liste des villes qui ne sont pas liés au canton à modifier
    requete1 = "SELECT comm_nom FROM COMM"
    requete1 = requete1 + " WHERE comm_insee NOT IN"
    requete1 = requete1 + " (SELECT comm_insee FROM CACO"
    requete1 = requete1 + " WHERE cant_num=" & frm_CommunesCanton.Txt_Code & ")"
      
    frm_CommunesCanton.Data_CC1.RecordSource = requete1
    frm_CommunesCanton.Data_CC1.Refresh
      
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not frm_CommunesCanton.Data_CC1.Recordset.EOF
      vntTemp = frm_CommunesCanton.Data_CC1.Recordset(0)
      If IsNull(vntTemp) Then vntTemp = ""
      frm_CommunesCanton.Lst_Lis_Can.AddItem CStr(vntTemp)
      frm_CommunesCanton.Data_CC1.Recordset.MoveNext
    Loop

    ' liste des communes qui sont liés au canton
    requete2 = "SELECT comm_nom FROM COMM"
    requete2 = requete2 + " WHERE comm_insee IN"
    requete2 = requete2 + " (SELECT comm_insee FROM CACO"
    requete2 = requete2 + " WHERE cant_num=" & frm_CommunesCanton.Txt_Code & ")"
      
    frm_CommunesCanton.Data_CC2.RecordSource = requete2
    frm_CommunesCanton.Data_CC2.Refresh
      
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not frm_CommunesCanton.Data_CC2.Recordset.EOF
      vntTemp = frm_CommunesCanton.Data_CC2.Recordset(0)
      If IsNull(vntTemp) Then vntTemp = ""
      frm_CommunesCanton.Lst_Lis_Can_Com.AddItem CStr(vntTemp)
      frm_CommunesCanton.Data_CC2.Recordset.MoveNext
    Loop
      
    frm_CommunesCanton.Show vbModal
  End If

End Sub

Private Sub cmd_Libelle_Click()

  Txt_Libelle = ""
  Txt_Libelle.SetFocus
End Sub

Private Sub cmd_Modifier_Click()

  If Not Dat_Canton.Recordset.EOF Then
    stb_ListeCanton.Panels(4).Text = "MODIFICATION"
    Load frm_CantonFiche
    frm_CantonFiche.Caption = "Modification d'un canton"
    frm_CantonFiche.Txt_code_canton.Enabled = False
    frm_CantonFiche.Txt_code_canton = Dat_Canton.Recordset("CANT_NUM")
    frm_CantonFiche.Txt_nom_canton = Dat_Canton.Recordset("CANT_NOM")
    
    frm_CantonFiche.Show vbModal
    stb_ListeCanton.Panels(4).Text = ""
    Dat_Canton.Refresh
  End If
  
End Sub

Private Sub cmd_Num_canton_Click()

  Txt_num_canton = ""
  Txt_num_canton.SetFocus
End Sub

Private Sub cmd_Creer_Click()

  stb_ListeCanton.Panels(4).Text = "CREATION"
  Load frm_CantonFiche
  frm_CantonFiche.Caption = "Création d'un canton"
  frm_CantonFiche.Txt_code_canton = ""
  frm_CantonFiche.Txt_nom_canton = ""
  frm_CantonFiche.Show vbModal

  stb_ListeCanton.Panels(4).Text = ""
  Dat_Canton.Refresh
  
End Sub

Private Sub cmd_Excel_Click()

  ExportExcel (Dat_Canton.RecordSource)

End Sub

Private Sub cmd_Quitter_Click()
    Unload Me
End Sub

Private Sub cmd_Supprimer_Click()

  If Not Dat_Canton.Recordset.EOF Then
    stb_ListeCanton.Panels(4).Text = "SUPPRESSION"
    Load frm_CantonFiche
    frm_CantonFiche.Caption = "Suppression d'un canton"
    frm_CantonFiche.Txt_code_canton.Enabled = False
    frm_CantonFiche.Txt_nom_canton.Enabled = False
    
    frm_CantonFiche.Txt_code_canton = Dat_Canton.Recordset("CANT_NUM")
    frm_CantonFiche.Txt_nom_canton = Dat_Canton.Recordset("CANT_NOM")
    
    frm_CantonFiche.Show vbModal
    stb_ListeCanton.Panels(4).Text = ""
    Dat_Canton.Refresh
  End If
  
End Sub

Private Sub cmd_Valider_Click()
    If Not Dat_Canton.Recordset.EOF Then
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
  Dat_Canton.DatabaseName = CheminBase
  Cmd_Actualiser_Click
End Sub

Private Sub grd_canton_HeadClick(ByVal ColIndex As Integer)

  Dim colonne As String
  Dim classem As String
  
  colonne = grd_canton.Columns(ColIndex).DataField
  Call Tri(Dat_Canton, grd_canton, colonne)
      
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_Canton.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_ListeCanton.Panels(4).Text = "Classement " & classem

End Sub

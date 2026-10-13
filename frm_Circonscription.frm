VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Circonscription 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liste des circonscriptions"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_Circonscription.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_Communes 
      Caption         =   "Associer des communes à la circonscription sélectionnée"
      Height          =   735
      Left            =   3720
      TabIndex        =   15
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Data Dat_Circonscriptions 
      Caption         =   "Dat_Circonscriptions"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   2895
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
         Picture         =   "frm_Circonscription.frx":0442
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
         Picture         =   "frm_Circonscription.frx":0544
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
         Picture         =   "frm_Circonscription.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Excel 
         Height          =   375
         Left            =   4560
         Picture         =   "frm_Circonscription.frx":0748
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Export EXCEL"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "frm_Circonscription.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprimer le régime d'assurance maladie sélectionné"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "frm_Circonscription.frx":09D4
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Créer un nouveau régime d'assurance maladie"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "frm_Circonscription.frx":0AD6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modifier le régime d'assurance maladie sélectionné"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   9255
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
         Caption         =   "&Nom"
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_code 
         Caption         =   "&Code"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.StatusBar stb_ListeCirconscriptions 
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
            TextSave        =   "09:25"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10557
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSDBGrid.DBGrid grd_Circonscriptions 
      Bindings        =   "frm_Circonscription.frx":0BD8
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "frm_Circonscription.frx":0BFB
      TabIndex        =   9
      Top             =   1560
      Width           =   9255
   End
End
Attribute VB_Name = "frm_Circonscription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Actualiser_Click()

Dim sel_code As String
Dim sel_nom As String
  
  If Txt_Code <> "" Then sel_code = " AND CIRC_NUM like """ & Txt_Code & "*"""
  If Txt_Nom <> "" Then sel_nom = " AND CIRC_NOM like """ & Txt_Nom & "*"""
      
  Dat_Circonscriptions.RecordSource = "SELECT * FROM CIRC WHERE TRUE" + _
                                      sel_code + _
                                      sel_nom + _
                                      ClauseOrder
      
  Me.MousePointer = vbHourglass
  Dat_Circonscriptions.Refresh
  Me.MousePointer = vbDefault
  
End Sub

Private Sub Cmd_canton_Click()

Dim requete1 As String
Dim requete2 As String
Dim requete3 As String
Dim requete4 As String

    If Not Dat_Circonscriptions.Recordset.EOF Then
        Load frm_CircoCantComm
                
        ' On rempli les textboxs avec les valeurs du canton sélectionné
        frm_CircoCantComm.Txt_Code_circo = Dat_Circonscriptions.Recordset("CIRC_NUM")
        frm_CircoCantComm.Txt_Nom_circo = Dat_Circonscriptions.Recordset("CIRC_NOM")
               
        ' liste des cantons qui ne sont pas liés à la circonscription à modifier
        requete1 = "SELECT DISTINCT(CANT_NOM) FROM CANT WHERE (CANT_NUM NOT IN"
        requete1 = requete1 + " (SELECT CANT_NUM FROM APPA"
        requete1 = requete1 + " WHERE CIRC_NUM=" & frm_CircoCantComm.Txt_Code_circo + "))"
        frm_CircoCantComm.Dat_Canton.RecordSource = requete1
        frm_CircoCantComm.Dat_Canton.Refresh
      
        ' remplissage du listbox avec les valeurs de la requète
        Do While Not frm_CircoCantComm.Dat_Canton.Recordset.EOF
          vntTemp = frm_CircoCantComm.Dat_Canton.Recordset(0)
          If IsNull(vntTemp) Then vntTemp = ""
          frm_CircoCantComm.Lst_Canton.AddItem CStr(vntTemp)
          frm_CircoCantComm.Dat_Canton.Recordset.MoveNext
        Loop
        
        ' liste des cantons qui sont liés à la circonscription à modifier
        requete2 = "SELECT DISTINCT(CANT_NOM) FROM CANT,APPA WHERE "
        requete2 = requete2 + "CANT.CANT_NUM=APPA.CANT_NUM AND CIRC_NUM=" & frm_CircoCantComm.Txt_Code_circo
        frm_CircoCantComm.Dat_CantonCirco.RecordSource = requete2
        frm_CircoCantComm.Dat_CantonCirco.Refresh
      
        ' remplissage du listbox avec les valeurs de la requète
        Do While Not frm_CircoCantComm.Dat_CantonCirco.Recordset.EOF
          vntTemp = frm_CircoCantComm.Dat_CantonCirco.Recordset(0)
          If IsNull(vntTemp) Then vntTemp = ""
          frm_CircoCantComm.Lst_CantonCirco.AddItem CStr(vntTemp)
          frm_CircoCantComm.Dat_CantonCirco.Recordset.MoveNext
        Loop
        
         ' liste des communes qui ne sont pas liés à la circonscription à modifier
        requete3 = "SELECT DISTINCT(COMM_NOM) FROM COMM WHERE (COMM_INSEE NOT IN"
        requete3 = requete3 + " (SELECT COMM_INSEE FROM APPA,CANT"
        requete3 = requete3 + " WHERE CIRC_NUM=" & frm_CircoCantComm.Txt_Code_circo + "))"
        frm_CircoCantComm.Dat_Commune.RecordSource = requete3
        frm_CircoCantComm.Dat_Commune.Refresh
      
        ' remplissage du listbox avec les valeurs de la requète
        Do While Not frm_CircoCantComm.Dat_Commune.Recordset.EOF
          vntTemp = frm_CircoCantComm.Dat_Commune.Recordset(0)
          If IsNull(vntTemp) Then vntTemp = ""
          frm_CircoCantComm.Lst_Commune.AddItem CStr(vntTemp)
          frm_CircoCantComm.Dat_Commune.Recordset.MoveNext
        Loop
        
        ' liste des communes qui sont liés à la circonscription à modifier
        requete4 = "SELECT DISTINCT(COMM_NOM) FROM COMM,APPA WHERE "
        requete4 = requete4 + "APPA.COMM_INSEE=COMM.COMM_INSEE AND "
        requete4 = requete4 + "APPA.CIRC_NUM=" & frm_CircoCantComm.Txt_Code_circo
        frm_CircoCantComm.Dat_CommuneCirco.RecordSource = requete4
        frm_CircoCantComm.Dat_CommuneCirco.Refresh
      
        ' remplissage du listbox avec les valeurs de la requète
        Do While Not frm_CircoCantComm.Dat_CommuneCirco.Recordset.EOF
          vntTemp = frm_CircoCantComm.Dat_CommuneCirco.Recordset(0)
          If IsNull(vntTemp) Then vntTemp = ""
          frm_CircoCantComm.Lst_CommuneCirco.AddItem CStr(vntTemp)
          frm_CircoCantComm.Dat_CommuneCirco.Recordset.MoveNext
        Loop
        
        frm_CircoCantComm.Show vbModal
    End If
    
End Sub


Private Sub cmd_code_Click()
    Txt_Code = ""
    Txt_Code.SetFocus
End Sub

Private Sub Cmd_Communes_Click()

  Dim reponse As String

  Dim vntTemp As Variant
  'Dim Code_Dep As String
  Dim requete1 As String
  Dim requete2 As String
  
  If Not Dat_Circonscriptions.Recordset.EOF Then
  
    Load Frm_CommunesCirco
          
    ' On rempli les textboxs avec les valeurs de la circonscription sélectionnée
    Frm_CommunesCirco.Txt_Code = Dat_Circonscriptions.Recordset("CIRC_NUM")
    Frm_CommunesCirco.Txt_Nom = Dat_Circonscriptions.Recordset("CIRC_NOM")
      
    ' data contenant les villes
    Frm_CommunesCirco.Data_Recherche.RecordSource = "SELECT * FROM COMM"
    Frm_CommunesCirco.Data_Recherche.Refresh
      
    ' liste des communes qui ne sont pas liées à la circonscription à modifier
    requete1 = "SELECT comm_nom FROM COMM"
    requete1 = requete1 + " WHERE comm_insee NOT IN"
    requete1 = requete1 + " (SELECT comm_insee FROM CICO"
    requete1 = requete1 + " WHERE circ_num=" & Frm_CommunesCirco.Txt_Code & ")"
      
    Frm_CommunesCirco.Data_CC1.RecordSource = requete1
    Frm_CommunesCirco.Data_CC1.Refresh
      
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Frm_CommunesCirco.Data_CC1.Recordset.EOF
      vntTemp = Frm_CommunesCirco.Data_CC1.Recordset(0)
      If IsNull(vntTemp) Then vntTemp = ""
      Frm_CommunesCirco.Lst_Lis_Can.AddItem CStr(vntTemp)
      Frm_CommunesCirco.Data_CC1.Recordset.MoveNext
    Loop

    ' liste des communes qui sont liés à la circonscription
    requete2 = "SELECT comm_nom FROM COMM"
    requete2 = requete2 + " WHERE comm_insee IN"
    requete2 = requete2 + " (SELECT comm_insee FROM CICO"
    requete2 = requete2 + " WHERE circ_num=" & Frm_CommunesCirco.Txt_Code & ")"
      
    Frm_CommunesCirco.Data_CC2.RecordSource = requete2
    Frm_CommunesCirco.Data_CC2.Refresh
      
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Frm_CommunesCirco.Data_CC2.Recordset.EOF
      vntTemp = Frm_CommunesCirco.Data_CC2.Recordset(0)
      If IsNull(vntTemp) Then vntTemp = ""
      Frm_CommunesCirco.Lst_Lis_Can_Com.AddItem CStr(vntTemp)
      Frm_CommunesCirco.Data_CC2.Recordset.MoveNext
    Loop
      
    Frm_CommunesCirco.Show vbModal
  End If

End Sub

Private Sub cmd_Creer_Click()

  stb_ListeCirconscriptions.Panels(4).Text = "CREATION"
  Load frm_CirconscriptionFiche
  With frm_CirconscriptionFiche
    .Caption = "Création d'une circonscription"
    .Txt_Code.Enabled = False
    .Txt_Code = NouveauCode("CIRC", "CIRC_NUM")
    .Txt_Nom = ""
    .Txt_adresse = ""
    .Txt_CP = ""
    .Txt_Ville = ""
    .Txt_responsable = ""
  End With
  
  frm_CirconscriptionFiche.Show vbModal

  stb_ListeCirconscriptions.Panels(4).Text = ""
  Dat_Circonscriptions.Refresh
  
End Sub

Private Sub cmd_Excel_Click()
    
  ExportExcel (Dat_Circonscriptions.RecordSource)
  
End Sub

Private Sub cmd_Modifier_Click()
    
  If Not Dat_Circonscriptions.Recordset.EOF Then
      stb_ListeCirconscriptions.Panels(4).Text = "MODIFICATION"
      Load frm_CirconscriptionFiche
      With frm_CirconscriptionFiche
        .Caption = "Modification d'une circonscription"
        .Txt_Code.Enabled = False
        .Txt_Code = Dat_Circonscriptions.Recordset("CIRC_NUM")
        .Txt_Nom = Dat_Circonscriptions.Recordset("CIRC_NOM")
        .Txt_adresse = IIf(IsNull(Dat_Circonscriptions.Recordset("CIRC_ADR")), "", Dat_Circonscriptions.Recordset("CIRC_ADR"))
        .Txt_CP = IIf(IsNull(Dat_Circonscriptions.Recordset("CIRC_CP")), "", Dat_Circonscriptions.Recordset("CIRC_CP"))
        .Txt_Ville = IIf(IsNull(Dat_Circonscriptions.Recordset("CIRC_VILLE")), "", Dat_Circonscriptions.Recordset("CIRC_VILLE"))
        .Txt_responsable = IIf(IsNull(Dat_Circonscriptions.Recordset("CIRC_RESP")), "", Dat_Circonscriptions.Recordset("CIRC_RESP"))
      End With
      frm_CirconscriptionFiche.Show vbModal
      stb_ListeCirconscriptions.Panels(4).Text = ""
      Dat_Circonscriptions.Refresh
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

  If Not Dat_Circonscriptions.Recordset.EOF Then
    stb_ListeCirconscriptions.Panels(4).Text = "SUPPRESSION"
    Load frm_CirconscriptionFiche
    With frm_CirconscriptionFiche
      .Caption = "Suppression d'une circonscription"
      .Txt_Code.Enabled = False
      .Txt_Nom.Enabled = False
      .Txt_adresse.Enabled = False
      .Txt_CP.Enabled = False
      .Txt_Ville.Enabled = False
      .Txt_responsable.Enabled = False
          
      .Txt_Code = Dat_Circonscriptions.Recordset("CIRC_NUM")
      .Txt_Nom = Dat_Circonscriptions.Recordset("CIRC_NOM")
      .Txt_adresse = IIf(IsNull(Dat_Circonscriptions.Recordset("CIRC_ADR")), "", Dat_Circonscriptions.Recordset("CIRC_ADR"))
      .Txt_CP = IIf(IsNull(Dat_Circonscriptions.Recordset("CIRC_CP")), "", Dat_Circonscriptions.Recordset("CIRC_CP"))
      .Txt_Ville = IIf(IsNull(Dat_Circonscriptions.Recordset("CIRC_VILLE")), "", Dat_Circonscriptions.Recordset("CIRC_VILLE"))
      .Txt_responsable = IIf(IsNull(Dat_Circonscriptions.Recordset("CIRC_RESP")), "", Dat_Circonscriptions.Recordset("CIRC_RESP"))
    End With
    
    frm_CirconscriptionFiche.Show vbModal
    stb_ListeCirconscriptions.Panels(4).Text = ""
    Dat_Circonscriptions.Refresh
  End If
  
End Sub

Private Sub cmd_Valider_Click()
    If Not Dat_Circonscriptions.Recordset.EOF Then
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Dat_Circonscriptions.DatabaseName = CheminBase
    Dat_Circonscriptions.RecordSource = "select * from circ"
    Cmd_Actualiser_Click
End Sub

Private Sub grd_Circonscriptions_headClick(ByVal ColIndex As Integer)

  Dim colonne As String
  Dim classem As String
  
  colonne = grd_Circonscriptions.Columns(ColIndex).DataField
  Call Tri(Dat_Circonscriptions, grd_Circonscriptions, colonne)
      
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_Circonscriptions.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_ListeCirconscriptions.Panels(4).Text = "Classement " & classem
    
End Sub

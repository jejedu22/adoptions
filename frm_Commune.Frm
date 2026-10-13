VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Commune 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liste des communes"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   1020
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_Commune.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6825
   ScaleMode       =   0  'User
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
         Picture         =   "frm_Commune.frx":0442
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
         Picture         =   "frm_Commune.frx":0544
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
         Picture         =   "frm_Commune.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Excel 
         Height          =   375
         Left            =   4560
         Picture         =   "frm_Commune.frx":0748
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Export EXCEL"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "frm_Commune.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprimer la commune sélectionnée"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "frm_Commune.frx":09D4
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Créer une nouvelle commune"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "frm_Commune.frx":0AD6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modifier les données de la commune sélectionnée"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Data Dat_Com 
      Caption         =   "Dat_Com"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM COMM ORDER BY COMM_CODEDEP, COMM_INSEE"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   9255
      Begin VB.TextBox Txt_Insee 
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   675
      End
      Begin VB.CommandButton cmd_Insee 
         Caption         =   "N° IN&SEE"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmd_Nom 
         Caption         =   "&Nom"
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Txt_Nom 
         Height          =   285
         Left            =   4680
         TabIndex        =   8
         Top             =   360
         Width           =   3795
      End
   End
   Begin MSDBGrid.DBGrid grd_com 
      Bindings        =   "frm_Commune.frx":0BD8
      Height          =   4815
      Left            =   1560
      OleObjectBlob   =   "frm_Commune.frx":0BEE
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1680
      Width           =   6135
   End
   Begin MSComctlLib.StatusBar stb_ListeCommune 
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
            TextSave        =   "09:26"
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
Attribute VB_Name = "frm_Commune"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Actualiser_Click()
  
  Dim sel_num_insee As String
  Dim sel_nom_com As String
  
  Me.MousePointer = vbHourglass
  If Txt_Insee <> "" Then sel_num_insee = " AND COMM_INSEE LIKE """ & Txt_Insee & "*"""
  If Txt_Nom <> "" Then sel_nom_com = " AND COMM_NOM LIKE """ & Txt_Nom & "*"""
  
  Dat_Com.RecordSource = "SELECT * FROM COMM WHERE TRUE" + _
                             sel_num_insee + _
                             sel_nom_com + _
                             " ORDER BY COMM_INSEE"
                             
  Dat_Com.Refresh
  Me.MousePointer = vbDefault
End Sub

Private Sub cmd_cantons_Click()

Dim requete1 As String
Dim requete2 As String
Dim requete3 As String
Dim requete4 As String

    If Not Dat_Com.Recordset.EOF Then
        Load Frm_CommCircoCant
                
        ' On rempli les textboxs avec les valeurs du canton sélectionné
        Frm_CommCircoCant.Txt_Code_commune = Dat_Com.Recordset("COMM_INSEE")
        Frm_CommCircoCant.Txt_Nom_commune = Dat_Com.Recordset("COMM_NOM")
               
        ' liste des circo qui ne sont pas liés à la commune à modifier
        requete1 = "SELECT DISTINCT(CIRC_NOM) FROM CIRC WHERE (CIRC_NUM NOT IN"
        requete1 = requete1 + " (SELECT CIRC_NUM FROM APPA"
        requete1 = requete1 + " WHERE COMM_INSEE=" & Frm_CommCircoCant.Txt_Code_commune + "))"
        Frm_CommCircoCant.Dat_Circo.RecordSource = requete1
        Frm_CommCircoCant.Dat_Circo.Refresh
      
        ' remplissage du listbox avec les valeurs de la requète
        Do While Not Frm_CommCircoCant.Dat_Circo.Recordset.EOF
          vntTemp = Frm_CommCircoCant.Dat_Circo.Recordset(0)
          If IsNull(vntTemp) Then vntTemp = ""
          Frm_CommCircoCant.Lst_Circo.AddItem CStr(vntTemp)
          Frm_CommCircoCant.Dat_Circo.Recordset.MoveNext
        Loop
        
        ' liste des circo qui sont liés à la commune à modifier
        requete2 = "SELECT DISTINCT(CIRC_NOM) FROM CIRC,APPA WHERE "
        requete2 = requete2 + "CIRC.CIRC_NUM=APPA.CIRC_NUM AND COMM_INSEE=" & Frm_CommCircoCant.Txt_Code_commune
        Frm_CommCircoCant.Dat_CircoComm.RecordSource = requete2
        Frm_CommCircoCant.Dat_CircoComm.Refresh
      
        ' remplissage du listbox avec les valeurs de la requète
        Do While Not Frm_CommCircoCant.Dat_CircoComm.Recordset.EOF
          vntTemp = Frm_CommCircoCant.Dat_CircoComm.Recordset(0)
          If IsNull(vntTemp) Then vntTemp = ""
          Frm_CommCircoCant.Lst_CircoCommune.AddItem CStr(vntTemp)
          Frm_CommCircoCant.Dat_CircoComm.Recordset.MoveNext
        Loop
        
         ' liste des cantons qui ne sont pas liés à la commune à modifier
        requete3 = "SELECT DISTINCT(CANT_NOM) FROM CANT WHERE (CANT_NUM NOT IN"
        requete3 = requete3 + " (SELECT CANT_NUM FROM APPA"
        requete3 = requete3 + " WHERE COMM_INSEE=" & Frm_CommCircoCant.Txt_Code_commune + "))"
        Frm_CommCircoCant.Dat_Cant.RecordSource = requete3
        Frm_CommCircoCant.Dat_Cant.Refresh
      
        ' remplissage du listbox avec les valeurs de la requète
        Do While Not Frm_CommCircoCant.Dat_Cant.Recordset.EOF
          vntTemp = Frm_CommCircoCant.Dat_Cant.Recordset(0)
          If IsNull(vntTemp) Then vntTemp = ""
          Frm_CommCircoCant.Lst_Cantons.AddItem CStr(vntTemp)
          Frm_CommCircoCant.Dat_Cant.Recordset.MoveNext
        Loop
        
        ' liste des cantons qui sont liés à la commune à modifier
        requete4 = "SELECT DISTINCT(CANT_NOM) FROM CANT,APPA WHERE "
        requete4 = requete4 + "APPA.CANT_NUM=CANT.CANT_NUM AND "
        requete4 = requete4 + "APPA.COMM_INSEE=" & Frm_CommCircoCant.Txt_Code_commune
        Frm_CommCircoCant.Dat_CantComm.RecordSource = requete4
        Frm_CommCircoCant.Dat_CantComm.Refresh
      
        ' remplissage du listbox avec les valeurs de la requète
        Do While Not Frm_CommCircoCant.Dat_CantComm.Recordset.EOF
          vntTemp = Frm_CommCircoCant.Dat_CantComm.Recordset(0)
          If IsNull(vntTemp) Then vntTemp = ""
          Frm_CommCircoCant.Lst_CantComm.AddItem CStr(vntTemp)
          Frm_CommCircoCant.Dat_CantComm.Recordset.MoveNext
        Loop
        
        Frm_CommCircoCant.Show vbModal
    End If
    
End Sub

Private Sub cmd_Creer_Click()

  stb_ListeCommune.Panels(4).Text = "CREATION"
  Load frm_CommuneFiche
  frm_CommuneFiche.Caption = "Création d'une commune"
  frm_CommuneFiche.Txt_num_insee = ""
  frm_CommuneFiche.Txt_nom_com = ""
  frm_CommuneFiche.Txt_cp_com = ""
  frm_CommuneFiche.Show vbModal

  stb_ListeCommune.Panels(4).Text = ""
  Dat_Com.Refresh

End Sub

Private Sub cmd_dept_Click()
  Txt_Depart = ""
  Txt_Depart.SetFocus
End Sub

Private Sub cmd_Excel_Click()

  ExportExcel (Dat_Com.RecordSource)
  
End Sub

Private Sub cmd_Insee_Click()
  Txt_Insee = ""
  Txt_Insee.SetFocus
End Sub

Private Sub cmd_nom_Click()
  Txt_Nom = ""
  Txt_Nom.SetFocus
End Sub

Private Sub cmd_Quitter_Click()
    Unload Me
End Sub

Private Sub cmd_Supprimer_Click()
    
  If Not Dat_Com.Recordset.EOF Then
    stb_ListeCommune.Panels(4).Text = "SUPPRESSION"
    Load frm_CommuneFiche
    frm_CommuneFiche.Caption = "Suppression d'une commune"
    frm_CommuneFiche.Txt_cp_com.Enabled = False
    frm_CommuneFiche.Txt_nom_com.Enabled = False
    frm_CommuneFiche.Txt_num_insee.Enabled = False
    frm_CommuneFiche.Txt_num_insee = Dat_Com.Recordset("COMM_INSEE")
    frm_CommuneFiche.Txt_nom_com = Dat_Com.Recordset("COMM_NOM")
    frm_CommuneFiche.Txt_cp_com = Dat_Com.Recordset("COMM_CP")
    
    frm_CommuneFiche.Show vbModal
    stb_ListeCommune.Panels(4).Text = ""
    Dat_Com.Refresh
  End If
End Sub

Private Sub cmd_Valider_Click()
    If Not Dat_Com.Recordset.EOF Then
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Dat_Com.DatabaseName = CheminBase
    Cmd_Actualiser_Click
End Sub

Private Sub cmd_Modifier_Click()

  If Not Dat_Com.Recordset.EOF Then
    stb_ListeCommune.Panels(4).Text = "MODIFICATION"
    Load frm_CommuneFiche
    frm_CommuneFiche.Caption = "Modification d'une commune"
    frm_CommuneFiche.Txt_num_insee.Enabled = False
    frm_CommuneFiche.Txt_num_insee = Dat_Com.Recordset("COMM_INSEE")
    frm_CommuneFiche.Txt_nom_com = Dat_Com.Recordset("COMM_NOM")
    frm_CommuneFiche.Txt_cp_com = Dat_Com.Recordset("COMM_CP")
    frm_CommuneFiche.Show vbModal
    
    stb_ListeCommune.Panels(4).Text = ""
    Dat_Com.Refresh
  End If
  
End Sub

Private Sub grd_com_headClick(ByVal ColIndex As Integer)

  Dim colonne As String
  Dim classem As String
  
  colonne = grd_com.Columns(ColIndex).DataField
  Call Tri(Dat_Com, grd_com, colonne)
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_Com.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_ListeCommune.Panels(4).Text = "Classement " & classem
  
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Commune_Canton 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Communes du canton"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   1020
   ClientWidth     =   9510
   Icon            =   "frm_Commune_Canton.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6120
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Annuler 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   7320
         Picture         =   "frm_Commune_Canton.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Annuler les modifications"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8760
         Picture         =   "frm_Commune_Canton.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8040
         Picture         =   "frm_Commune_Canton.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Enregistrer les modifications"
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.PictureBox GAUCHE 
      Height          =   495
      Left            =   4560
      Picture         =   "frm_Commune_Canton.frx":0748
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4200
      Width           =   615
   End
   Begin VB.PictureBox DROITE 
      Height          =   495
      Left            =   4560
      Picture         =   "frm_Commune_Canton.frx":0B8A
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2760
      Width           =   615
   End
   Begin VB.Data Dat_CC1 
      Caption         =   "Dat_CC1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Dat_CC2 
      Caption         =   "Dat_CC2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Dat_Recherche 
      Caption         =   "Dat_Recherche"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Txt_Nom_canton 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Txt_Code_canton 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Width           =   615
   End
   Begin VB.ListBox Lst_Lis_Can_Com 
      Height          =   3570
      ItemData        =   "frm_Commune_Canton.frx":0FCC
      Left            =   5640
      List            =   "frm_Commune_Canton.frx":0FCE
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1920
      Width           =   2895
   End
   Begin VB.ListBox Lst_Lis_Can 
      Height          =   3570
      ItemData        =   "frm_Commune_Canton.frx":0FD0
      Left            =   1200
      List            =   "frm_Commune_Canton.frx":0FD2
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Width           =   2895
   End
   Begin MSComctlLib.StatusBar stb_ComCant 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   5865
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
            TextSave        =   "15/07/99"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10:07"
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
   Begin VB.Label lbl_liste2 
      Caption         =   "Liste des communes appartenant au canton"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   12
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lbl_liste 
      Caption         =   "Liste des communes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lbl_nom 
      Caption         =   "Canton"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lbl_num 
      Caption         =   "N° Canton"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frm_Commune_Canton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Annuler_Click()

  Dim vntTemp As Variant
  Dim Code_Dep As String
  Dim requete1 As String
  Dim requete2 As String
  
  ' vidage des listes
  Do While Lst_Lis_Can.ListCount > 0
    Lst_Lis_Can.RemoveItem (0)
  Loop
  Do While Lst_Lis_Can_Com.ListCount > 0
    Lst_Lis_Can_Com.RemoveItem (0)
  Loop
  
  ' département de la commune
  Code_Dep = Left(CStr(frm_Canton.Dat_Canton.recordset("CANT_NUM")), 2)
  
  ' liste des communes qui ne sont pas liées au canton à modifier, mais qui sont du même département
  requete1 = "SELECT COMM_NOM FROM COMM WHERE COMM_INSEE NOT IN"
  requete1 = requete1 + " (SELECT COMM_INSEE FROM APPA"
  requete1 = requete1 + " WHERE CANT_NUM=" & frm_Commune_Canton.Txt_code_canton
  requete1 = requete1 + " AND COMM_INSEE LIKE """ & Code_Dep & "*"")"
  
  frm_Commune_Canton.dat_CC1.RecordSource = requete1
  frm_Commune_Canton.dat_CC1.Refresh
  
  ' remplissage du listbox avec les valeurs de la requète
  Do While Not frm_Commune_Canton.dat_CC1.recordset.EOF
    vntTemp = frm_Commune_Canton.dat_CC1.recordset(0)
    If IsNull(vntTemp) Then vntTemp = ""
    frm_Commune_Canton.Lst_Lis_Can.AddItem CStr(vntTemp)
    frm_Commune_Canton.dat_CC1.recordset.MoveNext
  Loop
  
  ' liste des communes qui sont liés au canton
  requete2 = "SELECT COMM_NOM FROM COMM"
  requete2 = requete2 + " WHERE COMM_INSEE IN"
  requete2 = requete2 + " (SELECT COMM_INSEE FROM APPA"
  requete2 = requete2 + " WHERE CANT_NUM=" & frm_Commune_Canton.Txt_code_canton & ")"
  
  frm_Commune_Canton.dat_CC2.RecordSource = requete2
  frm_Commune_Canton.dat_CC2.Refresh
  
  ' remplissage du listbox avec les valeurs de la requète
  Do While Not frm_Commune_Canton.dat_CC2.recordset.EOF
    vntTemp = frm_Commune_Canton.dat_CC2.recordset(0)
    If IsNull(vntTemp) Then vntTemp = ""
    frm_Commune_Canton.Lst_Lis_Can_Com.AddItem CStr(vntTemp)
    frm_Commune_Canton.dat_CC2.recordset.MoveNext
  Loop
  stb_ComCant.Panels(4).Text = "MAJ annulée"
  
End Sub

Private Sub cmd_Quitter_Click()
    Unload Me
End Sub

Private Sub DROITE_Click()
  Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite
  If Lst_Lis_Can.ListIndex > -1 Then
    ind = Lst_Lis_Can.ListIndex
    Lst_Lis_Can_Com.AddItem Lst_Lis_Can.List(ind)
    Lst_Lis_Can.RemoveItem ind
  End If
End Sub

Private Sub Form_Load()
  dat_CC1.DatabaseName = CheminBase
  dat_CC1.Refresh
  dat_CC2.DatabaseName = CheminBase
  dat_CC2.Refresh
  dat_Recherche.DatabaseName = CheminBase
  dat_Recherche.Refresh
End Sub

Private Sub GAUCHE_Click()
  Dim ind As Integer
  'Fait passer un élément de la liste de droite dans celle de gauche
  If Lst_Lis_Can_Com.ListIndex > -1 Then
    ind = Lst_Lis_Can_Com.ListIndex
    Lst_Lis_Can.AddItem Lst_Lis_Can_Com.List(ind)
    Lst_Lis_Can_Com.RemoveItem ind
  End If
End Sub


Private Sub Lst_Lis_Can_Com_DblClick()
  Dim ind As Integer
  'Fait passer un élément de la liste de droite dans celle de gauche
  If Lst_Lis_Can_Com.ListIndex > -1 Then
    ind = Lst_Lis_Can_Com.ListIndex
    Lst_Lis_Can.AddItem Lst_Lis_Can_Com.List(ind)
    Lst_Lis_Can_Com.RemoveItem ind
  End If
End Sub

Private Sub Lst_Lis_Can_Com_GotFocus()
  stb_ComCant.Panels(3).Text = "Double clic ou entrée fait passer dans l'autre liste"
End Sub

Private Sub Lst_Lis_Can_Com_KeyPress(KeyAscii As Integer)
  Dim ind As Integer
  'Fait passer un élément de la liste de droite dans celle de gauche
  If KeyAscii = vbKeyReturn Then
    If Lst_Lis_Can_Com.ListIndex > -1 Then
      ind = Lst_Lis_Can_Com.ListIndex
      Lst_Lis_Can.AddItem Lst_Lis_Can_Com.List(ind)
      Lst_Lis_Can_Com.RemoveItem ind
    End If
  End If
End Sub

Private Sub Lst_Lis_Can_Com_LostFocus()
  stb_ComCant.Panels(3).Text = ""
End Sub

Private Sub Lst_Lis_Can_DblClick()
  'Fait passer un élément de la liste de gauche dans celle de droite
  Dim ind As Integer
  If Lst_Lis_Can.ListIndex > -1 Then
    ind = Lst_Lis_Can.ListIndex
    Lst_Lis_Can_Com.AddItem Lst_Lis_Can.List(ind)
    Lst_Lis_Can.RemoveItem ind
  End If
End Sub

Private Sub Lst_Lis_Can_GotFocus()
    stb_ComCant.Panels(3).Text = "Double clic ou entrée fait passer dans l'autre liste"
End Sub

Private Sub Lst_Lis_Can_KeyPress(KeyAscii As Integer)
  Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite
  If KeyAscii = vbKeyReturn Then
    If Lst_Lis_Can.ListIndex > -1 Then
      ind = Lst_Lis_Can.ListIndex
      Lst_Lis_Can_Com.AddItem Lst_Lis_Can.List(ind)
      Lst_Lis_Can.RemoveItem ind
    End If
  End If
End Sub

Private Sub Lst_Lis_Can_LostFocus()
    stb_ComCant.Panels(3).Text = ""
End Sub

Private Sub cmd_Valider_Click()

  Dim reponse As String
  Dim requete As String
  Dim Code_Dep As String
  Dim i As Integer
  Dim selection As String
  Dim num_ins As Integer
  
  'Permet la validation des nouvelles liaisons entre un canton et ses villes
  reponse = MsgBox("Etes-vous sur de vouloir enregistrer les modifications ?", vbYesNo + vbQuestion + vbApplicationModal)
  If reponse = vbYes Then
    Code_Dep = CStr(frm_Canton.Dat_Canton.recordset("CANT_CODEDEP"))
    requete = "DELETE FROM APPA"
    requete = requete + " WHERE CANT_NUM=" & Txt_code_canton
    dat_CC2.Database.Execute requete
    For i = 0 To Lst_Lis_Can_Com.ListCount - 1
      selection = Lst_Lis_Can_Com.List(i)
      dat_Recherche.recordset.FindFirst "COMM_NOM=""" & selection & """"
      num_ins = dat_Recherche.recordset(0)
      
      requete = "INSERT INTO APPA(CANT_NUM,COMM_INSEE)"
      requete = requete + " VALUES (""" & Txt_code_canton & """,""" & num_ins & """);"
      
      dat_CC2.DatabaseName = CheminBase
      dat_CC2.Database.Execute requete
      dat_CC2.Refresh
    Next i
    dat_CC2.Database.Execute requete
    stb_ComCant.Panels(4).Text = "MAJ effectuée"
  Else
    stb_ComCant.Panels(4).Text = "MAJ refusée"
  End If
End Sub


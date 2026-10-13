VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Canton_Commune 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Séparation d'une ville en plusieurs cantons"
   ClientHeight    =   6120
   ClientLeft      =   855
   ClientTop       =   1005
   ClientWidth     =   9510
   ForeColor       =   &H00000000&
   Icon            =   "frm_Canton_Commune.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6120
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8040
         Picture         =   "frm_Canton_Commune.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Enregistrer les modifications"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8760
         Picture         =   "frm_Canton_Commune.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmd_Annuler 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   7320
         Picture         =   "frm_Canton_Commune.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Annuler les modifications"
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Data dat_Recherche 
      Caption         =   "dat_Recherche"
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
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dat_CC2 
      Caption         =   "dat_CC2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data dat_CC1 
      Caption         =   "dat_CC1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox Lst_Lis_Can_Com 
      Height          =   3570
      ItemData        =   "frm_Canton_Commune.frx":0748
      Left            =   5640
      List            =   "frm_Canton_Commune.frx":074A
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox Txt_ville 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Txt_num_insee 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox pic_droite 
      Height          =   495
      Left            =   4560
      Picture         =   "frm_Canton_Commune.frx":074C
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2760
      Width           =   615
   End
   Begin VB.PictureBox pic_gauche 
      Height          =   495
      Left            =   4560
      Picture         =   "frm_Canton_Commune.frx":0B8E
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4200
      Width           =   615
   End
   Begin MSComctlLib.StatusBar stb_CantCom 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
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
            TextSave        =   "10/08/99"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15:56"
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
   Begin VB.ListBox Lst_Lis_Can 
      Height          =   3570
      ItemData        =   "frm_Canton_Commune.frx":0FD0
      Left            =   1200
      List            =   "frm_Canton_Commune.frx":0FD2
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Liste des cantons dont  fait partie la commune"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Liste des cantons"
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
      Left            =   1800
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Ville"
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
      Left            =   4800
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "N° INSEE"
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
      Left            =   2520
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frm_Canton_Commune"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Quitter_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  dat_CC1.DatabaseName = CheminBase
  dat_CC1.Refresh
  dat_CC2.DatabaseName = CheminBase
  dat_CC2.Refresh
  dat_Recherche.DatabaseName = CheminBase
  dat_Recherche.Refresh
End Sub

Private Sub cmd_Annuler_Click()
  Dim vntTemp As Variant
  Dim requete1 As String
  Dim requete2 As String
  Dim Code_Dep As String
  
  Code_Dep = CStr(frm_Commune.Dat_Com.recordset("COMM_CODEDEP"))
  
  ' vidage des listes
  
  Do While Lst_Lis_Can.ListCount > 0
    Lst_Lis_Can.RemoveItem (0)
  Loop
  
  Do While Lst_Lis_Can_Com.ListCount > 0
    Lst_Lis_Can_Com.RemoveItem (0)
  Loop
  
  ' liste des cantons qui ne sont pas liés à la commune à modifier
  
  requete1 = "SELECT CANT_NOM FROM CANT"
  requete1 = requete1 + " WHERE CANT_CODEDEP = " & Code_Dep
  requete1 = requete1 + " AND CANT_NUM NOT IN"
  requete1 = requete1 + " (SELECT CANT_NUM FROM APPA"
  requete1 = requete1 + " WHERE COMM_CODEDEP=" & Code_Dep
  requete1 = requete1 + " AND COMM_INSEE = " & frm_Canton_Commune.Txt_num_insee & ")"
  frm_Canton_Commune.dat_CC1.RecordSource = requete1
  frm_Canton_Commune.dat_CC1.Refresh
  
  ' remplissage du listbox avec les valeurs de la requète
  
  Do While Not frm_Canton_Commune.dat_CC1.recordset.EOF
    vntTemp = frm_Canton_Commune.dat_CC1.recordset(0)
    If IsNull(vntTemp) Then vntTemp = ""
    frm_Canton_Commune.Lst_Lis_Can.AddItem CStr(vntTemp)
    frm_Canton_Commune.dat_CC1.recordset.MoveNext
  Loop
  
  ' liste des cantons qui sont liés à la commune
  
  requete2 = "SELECT CANT_NOM FROM CANT"
  requete2 = requete2 + " WHERE CANT_CODEDEP = " & Code_Dep
  requete2 = requete2 + " AND CANT_NUM IN"
  requete2 = requete2 + " (SELECT CANT_NUM FROM APPA"
  requete2 = requete2 + " WHERE COMM_CODEDEP=" & Code_Dep
  requete2 = requete2 + " AND COMM_INSEE = " & frm_Canton_Commune.Txt_num_insee & ")"
  frm_Canton_Commune.dat_CC2.RecordSource = requete2
  frm_Canton_Commune.dat_CC2.Refresh
  
  ' remplissage du listbox avec les valeurs de la requète
  
  Do While Not frm_Canton_Commune.dat_CC2.recordset.EOF
    vntTemp = frm_Canton_Commune.dat_CC2.recordset(0)
    If IsNull(vntTemp) Then vntTemp = ""
    frm_Canton_Commune.Lst_Lis_Can_Com.AddItem CStr(vntTemp)
    frm_Canton_Commune.dat_CC2.recordset.MoveNext
  Loop
  
  stb_CantCom.Panels(4).Text = "MAJ annulée"

End Sub

Private Sub cmd_Valider_Click()
  
  Dim reponse As String
  Dim requete As String
  Dim i As Integer
  Dim selection As String
  Dim int_code_canton As Integer
  Dim Code_Dep As String
  
  Code_Dep = CStr(frm_Commune.Dat_Com.recordset("COMM_CODEDEP"))
  
  'Enregistre dans la base de données (table appartient) les relations entre la ville et son(ses) canton(s)
  reponse = MsgBox("Etes-vous sur de vouloir enregistrer les modifications ?", vbYesNo + vbQuestion + vbApplicationModal)
  If reponse = vbYes Then
    requete = "DELETE FROM APPA"
    requete = requete + " WHERE COMM_CODEDEP=" & Code_Dep
    requete = requete + " AND COMM_INSEE = " & Txt_num_insee
    dat_CC2.Database.Execute requete
    
    For i = 0 To Lst_Lis_Can_Com.ListCount - 1
      selection = Lst_Lis_Can_Com.List(i)
      dat_Recherche.recordset.FindFirst "CANT_NOM='" & selection & "'"
      int_code_canton = dat_Recherche.recordset(0)
      
      requete = "INSERT INTO APPA(COMM_CODEDEP,CANT_NUM,COMM_INSEE)"
      requete = requete + " VALUES (" & Code_Dep & "," & int_code_canton & "," & Txt_num_insee & ")"
      
      dat_CC2.Database.Execute requete
      dat_CC2.Refresh
    Next i
    dat_CC2.Database.Execute requete
    stb_CantCom.Panels(4).Text = "MAJ effectuée"
  Else
    stb_CantCom.Panels(4).Text = "MAJ refusée"
  End If
End Sub

Private Sub Lst_Lis_Can_Com_DblClick()
  Dim ind As Integer
  'Déplace un élément de la liste de droite vers la liste de gauche
  If Lst_Lis_Can_Com.ListIndex > -1 Then
    ind = Lst_Lis_Can_Com.ListIndex
    Lst_Lis_Can.AddItem Lst_Lis_Can_Com.List(ind)
    Lst_Lis_Can_Com.RemoveItem ind
  End If
End Sub

Private Sub Lst_Lis_Can_Com_GotFocus()
  stb_CantCom.Panels(3).Text = "Double clic ou entrée fait passer dans l'autre liste"
End Sub

Private Sub Lst_Lis_Can_Com_KeyPress(KeyAscii As Integer)
  Dim ind As Integer
  'Déplace un élément de la liste de droite vers la liste de gauche
  If KeyAscii = vbKeyReturn Then
    If Lst_Lis_Can_Com.ListIndex > -1 Then
      ind = Lst_Lis_Can_Com.ListIndex
      Lst_Lis_Can.AddItem Lst_Lis_Can_Com.List(ind)
      Lst_Lis_Can_Com.RemoveItem ind
    End If
  End If
End Sub

Private Sub Lst_Lis_Can_Com_LostFocus()
  stb_CantCom.Panels(3).Text = ""
End Sub

Private Sub Lst_Lis_Can_DblClick()
  Dim ind As Integer
  'Déplace un élément de la liste de gauche vers la liste de droite
  If Lst_Lis_Can.ListIndex > -1 Then
    ind = Lst_Lis_Can.ListIndex
    Lst_Lis_Can_Com.AddItem Lst_Lis_Can.List(ind)
    Lst_Lis_Can.RemoveItem ind
  End If
End Sub

Private Sub Lst_Lis_Can_GotFocus()
  stb_CantCom.Panels(3).Text = "Double clic ou entrée fait passer dans l'autre liste"
End Sub

Private Sub Lst_Lis_Can_KeyPress(KeyAscii As Integer)
  Dim ind As Integer
  'Déplace un élément de la liste de gauche vers la liste de droite
  If KeyAscii = vbKeyReturn Then
    If Lst_Lis_Can.ListIndex > -1 Then
      ind = Lst_Lis_Can.ListIndex
      Lst_Lis_Can_Com.AddItem Lst_Lis_Can.List(ind)
      Lst_Lis_Can.RemoveItem ind
    End If
  End If
End Sub

Private Sub Lst_Lis_Can_LostFocus()
  stb_CantCom.Panels(3).Text = ""
End Sub

Private Sub pic_droite_Click()
  Dim ind As Integer
  'Déplace un élément de la liste de gauche vers la liste de droite
  If Lst_Lis_Can.ListIndex > -1 Then
    ind = Lst_Lis_Can.ListIndex
    Lst_Lis_Can_Com.AddItem Lst_Lis_Can.List(ind)
    Lst_Lis_Can.RemoveItem ind
  End If
End Sub

Private Sub pic_gauche_Click()
  Dim ind As Integer
  'Déplace un élément de la liste de droite vers la liste de gauche
  If Lst_Lis_Can_Com.ListIndex > -1 Then
    ind = Lst_Lis_Can_Com.ListIndex
    Lst_Lis_Can.AddItem Lst_Lis_Can_Com.List(ind)
    Lst_Lis_Can_Com.RemoveItem ind
  End If
End Sub



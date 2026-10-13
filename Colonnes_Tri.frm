VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Colonnes_Tri 
   Caption         =   "Choix de l'ordre"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   ControlBox      =   0   'False
   Icon            =   "Colonnes_Tri.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8775
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox Lst_Sens 
      Height          =   3765
      ItemData        =   "Colonnes_Tri.frx":0442
      Left            =   7080
      List            =   "Colonnes_Tri.frx":0444
      TabIndex        =   12
      ToolTipText     =   "Double cliquer sur la liste pour changer le sens du tri"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Cmd_Descendre 
      Height          =   495
      Left            =   3480
      Picture         =   "Colonnes_Tri.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Cmd_Droite 
      Height          =   495
      Left            =   3480
      Picture         =   "Colonnes_Tri.frx":0888
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Cmd_Gauche 
      Height          =   495
      Left            =   3480
      Picture         =   "Colonnes_Tri.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Cmd_Monter 
      Height          =   495
      Left            =   3480
      Picture         =   "Colonnes_Tri.frx":110C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   615
   End
   Begin VB.Data Dat_woby 
      Caption         =   "Dat_woby"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Dat_wsel 
      Caption         =   "Dat_wsel"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8535
      Begin VB.Data Dat_Interim 
         Caption         =   "Dat_interim"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   7800
         Picture         =   "Colonnes_Tri.frx":154E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Annuler 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   7080
         Picture         =   "Colonnes_Tri.frx":1650
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_ChoixColonnes 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5550
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "18/05/2010"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "14:30"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7779
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin VB.ListBox Lst_ColonneTri 
      Height          =   3765
      ItemData        =   "Colonnes_Tri.frx":1752
      Left            =   4320
      List            =   "Colonnes_Tri.frx":1754
      TabIndex        =   6
      Top             =   1440
      Width           =   2775
   End
   Begin VB.ListBox Lst_colonnes 
      Height          =   3765
      ItemData        =   "Colonnes_Tri.frx":1756
      Left            =   360
      List            =   "Colonnes_Tri.frx":1758
      TabIndex        =   7
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Liste des colonnes ""triées"" :"
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
      Left            =   4320
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Liste des colonnes ""non triées"" :"
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
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
   End
End
Attribute VB_Name = "Colonnes_Tri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Annuler_Click()
  Unload Me
End Sub

Private Sub Cmd_Descendre_Click()

Dim ind1, ind2 As Integer
Dim Desc As String

  'Fait descendre un élément de la liste de droite

  If Lst_ColonneTri.ListIndex > -1 And Lst_ColonneTri.ListIndex < Lst_ColonneTri.ListCount - 1 Then
    ind1 = Lst_ColonneTri.ListIndex
    ind2 = ind1 + 1
    Desc = Lst_ColonneTri.List(ind1)
    Lst_ColonneTri.RemoveItem (ind1)
    Lst_ColonneTri.AddItem Desc, ind2
  End If

End Sub

Private Sub Cmd_Droite_Click()

Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

  If Lst_colonnes.ListIndex > -1 Then
    ind = Lst_colonnes.ListIndex
    Lst_ColonneTri.AddItem Lst_colonnes.List(ind)
    Lst_colonnes.RemoveItem ind
    Lst_Sens.AddItem "Ascendant"
  End If
  
End Sub

Private Sub Cmd_Gauche_Click()

Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

  If Lst_ColonneTri.ListIndex > -1 Then
    ind = Lst_ColonneTri.ListIndex
    Lst_colonnes.AddItem Lst_ColonneTri.List(ind)
    Lst_ColonneTri.RemoveItem ind
    Lst_Sens.RemoveItem ind
  End If
  
End Sub

Private Sub Cmd_Monter_Click()

Dim ind1, ind2 As Integer
Dim Mont As String

  'Fait monter un élément de la liste de droite

  If Lst_ColonneTri.ListIndex > 0 Then
    ind2 = Lst_ColonneTri.ListIndex
    ind1 = ind2 - 1
    Mont = Lst_ColonneTri.List(ind2)
    Lst_ColonneTri.RemoveItem (ind2)
    Lst_ColonneTri.AddItem Mont, ind1
    
  End If
  
End Sub


Private Sub cmd_Valider_Click()
  
Dim sql As String
Dim i As Integer
Dim bloque As String

'On Error GoTo Rupture_Integrite

  'supprime les tris existants pour la source de données
  sql = "delete * from woby where wsou_id=" + CStr(SourcesDonnées_fiche.Txt_id)
  Workspaces(0).Databases(0).Execute sql, dbFailOnError
  'crée les nouveaux tris

  For i = 0 To Lst_ColonneTri.ListCount - 1
    If Lst_Sens.List(i) = "Ascendant" Then
      bloque = "A"
    Else
      bloque = "D"
    End If
    sql = "select wcol_nom from wsel where wsou_id=" + CStr(SourcesDonnées_fiche.Txt_id)
    sql = sql + " and wsel_alias=" + Iifch(Lst_ColonneTri.List(i))
    With Dat_Interim
      .RecordSource = sql
      .Refresh
      .Recordset.MoveFirst
      .Recordset.MoveLast
    End With
    If Dat_Interim.Recordset.RecordCount = 1 Then
      'modifié le 18/05/2010 par Caroline (précision sur l'ordre des champs)
      sql = "insert into woby (WSOU_ID,WCOL_NOM,WOBY_NUMORDRE,WOBY_SENS) values(" + CStr(SourcesDonnées_fiche.Txt_id)
      sql = sql + "," + Iifch(Dat_Interim.Recordset("wcol_nom"))
      sql = sql + "," + CStr(i) + "," + Iifch(bloque) + ")"
      Workspaces(0).Databases(0).Execute sql, dbFailOnError
    End If
  Next i
  
  
  Unload Me
End Sub


Private Sub Command2_Click()

End Sub

Private Sub DROITE_Click()

Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

  If Lst_colonnes.ListIndex > -1 Then
    ind = Lst_colonnes.ListIndex
    Lst_ColonneTri.AddItem Lst_colonnes.List(ind)
    Lst_colonnes.RemoveItem ind
  End If
    
End Sub


Private Sub Form_Load()

  Dat_wsel.DatabaseName = CheminBase
  Dat_woby.DatabaseName = CheminBase
  Dat_Interim.DatabaseName = CheminBase
  
End Sub


Private Sub GAUCHE_Click()

Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

  If Lst_ColonneTri.ListIndex > -1 Then
    ind = Lst_ColonneTri.ListIndex
    Lst_colonnes.AddItem Lst_ColonneTri.List(ind)
    Lst_ColonneTri.RemoveItem ind
  End If
  
End Sub

Private Sub Lst_colonnes_DblClick()

Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

  If Lst_colonnes.ListIndex > -1 Then
    ind = Lst_colonnes.ListIndex
    Lst_ColonneTri.AddItem Lst_colonnes.List(ind)
    Lst_colonnes.RemoveItem ind
    Lst_Sens.AddItem "Ascendant"
  End If
  
End Sub


Private Sub Lst_ColonneTri_DblClick()

Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

  If Lst_ColonneTri.ListIndex > -1 Then
    ind = Lst_ColonneTri.ListIndex
    Lst_colonnes.AddItem Lst_ColonneTri.List(ind)
    Lst_ColonneTri.RemoveItem ind
    Lst_Sens.RemoveItem ind
  End If

End Sub


Private Sub Lst_Sens_DblClick()

Dim ind As Integer
  'Change le sens du tri

  If Lst_Sens.ListIndex > -1 Then
    ind = Lst_Sens.ListIndex
    If Lst_Sens.List(ind) = "Ascendant" Then
      Lst_Sens.List(ind) = "Descendant"
    Else
      Lst_Sens.List(ind) = "Ascendant"
    End If
  End If

End Sub



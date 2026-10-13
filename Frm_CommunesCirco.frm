VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_CommunesCirco 
   Caption         =   "Communes de la circonscription"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   ControlBox      =   0   'False
   Icon            =   "Frm_CommunesCirco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Annuler 
         Default         =   -1  'True
         Height          =   375
         Left            =   7800
         Picture         =   "Frm_CommunesCirco.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8520
         Picture         =   "Frm_CommunesCirco.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9255
      Begin VB.PictureBox GAUCHE 
         Height          =   495
         Left            =   4200
         Picture         =   "Frm_CommunesCirco.frx":0646
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3480
         Width           =   615
      End
      Begin VB.PictureBox DROITE 
         Height          =   495
         Left            =   4200
         Picture         =   "Frm_CommunesCirco.frx":0A88
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2040
         Width           =   615
      End
      Begin VB.Data Data_CC1 
         Caption         =   "Data_CC1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data Data_CC2 
         Caption         =   "Data_CC2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2880
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Data Data_Recherche 
         Caption         =   "Data_Recherche"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox Txt_Nom 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Txt_Code 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox Lst_Lis_Can 
         Height          =   3570
         ItemData        =   "Frm_CommunesCirco.frx":0ECA
         Left            =   840
         List            =   "Frm_CommunesCirco.frx":0ECC
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   2895
      End
      Begin VB.ListBox Lst_Lis_Can_Com 
         Height          =   3570
         ItemData        =   "Frm_CommunesCirco.frx":0ECE
         Left            =   5280
         List            =   "Frm_CommunesCirco.frx":0ED0
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Liste des communes appartenant à la circonscription"
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
         Left            =   5400
         TabIndex        =   10
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label3 
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
         Left            =   1200
         TabIndex        =   9
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Circonscription"
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
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "N° Circonscription"
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
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar stb_ListeCanton 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   5610
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
            TextSave        =   "23/11/99"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10:23"
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
Attribute VB_Name = "Frm_CommunesCirco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Annuler_Click()
  Unload Me
End Sub

Private Sub cmd_Valider_Click()
  
  Dim reponse As String
  Dim requete As String
  Dim Code_Dep As String
  Dim i As Integer
  Dim selection As String
  Dim num_ins As Integer
  
  'Permet la validation des nouvelles liaisons entre une circonscription et ses communes
  
  reponse = MsgBox("Etes-vous sur de vouloir enregistrer les modifications ?", vbYesNo + vbQuestion + vbApplicationModal)
  
  If reponse = vbYes Then
    requete = "DELETE * FROM CICO WHERE circ_num=" & Txt_Code
    Data_CC2.Database.Execute requete
    
    For i = 0 To Lst_Lis_Can_Com.ListCount - 1
      selection = Lst_Lis_Can_Com.List(i)
      Data_Recherche.Recordset.FindFirst "comm_nom=""" & selection & """"
      num_ins = Data_Recherche.Recordset(0)
      
      requete = "INSERT INTO CICO VALUES (""" & num_ins & """,""" & Txt_Code & """);"
      Data_CC2.Database.Execute requete
      Data_CC2.Refresh
    Next i
    
    Data_CC2.Database.Execute requete
    stb_ListeCanton.Panels(3).Text = "MAJ effectuée"
    Unload Me
  Else
    stb_ListeCanton.Panels(3).Text = "MAJ refusée"
  End If
  
  Exit Sub
  
Rupture_Integrite:
'Affiche le code d'erreur approprié
    Select Case Err
        Case 13
            MsgBox "Valeur incompatible avec le type de colonne", vbOKOnly + vbExclamation
      
        Case 3022
            MsgBox "Doublon, création impossible", vbOKOnly + vbExclamation
      
        Case 3200
            MsgBox "Reference, suppression impossible", vbOKOnly + vbExclamation
      
        Case 3075
            MsgBox "L'un des champs contient une cote, mise à jour impossible", vbOKOnly + vbExclamation
      
        Case 3315
            MsgBox "Un champ obligatoire n'est pas renseigné, mise à jour impossible", vbOKOnly + vbExclamation
      
        Case Else
            MsgBox "Erreur" + CStr(Err) + " " + Err.Description + " survenue, mise à jour impossible", vbOKOnly + vbExclamation
    End Select
    
    Exit Sub
  
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

  Data_CC1.DatabaseName = CheminBase
  Data_CC2.DatabaseName = CheminBase
  Data_Recherche.DatabaseName = CheminBase
  
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


Private Sub Lst_Lis_Can_DblClick()

  Dim ind As Integer
  
  'Fait passer un élément de la liste de gauche dans celle de droite
  
  If Lst_Lis_Can.ListIndex > -1 Then
    
    ind = Lst_Lis_Can.ListIndex
    Lst_Lis_Can_Com.AddItem Lst_Lis_Can.List(ind)
    Lst_Lis_Can.RemoveItem ind
  
  End If

End Sub



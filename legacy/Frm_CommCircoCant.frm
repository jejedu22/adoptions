VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_CommCircoCant 
   Caption         =   "Association d'une commune aux circonscriptions et communes"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   Icon            =   "Frm_CommCircoCant.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txt_Canton 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      TabIndex        =   6
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Txt_Circo 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Txt_Code_commune 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Txt_Nom_commune 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   2535
   End
   Begin MSComctlLib.StatusBar stb_CantCircoComm 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6270
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "26/08/99"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15:27"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9913
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Circonscription"
      TabPicture(0)   =   "Frm_CommCircoCant.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_Canton"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra_OutilsCirco"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Canton"
      TabPicture(1)   =   "Frm_CommCircoCant.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_OutilsCommune"
      Tab(1).Control(1)=   "Fra_Commune"
      Tab(1).ControlCount=   2
      Begin VB.Frame Fra_OutilsCommune 
         Height          =   600
         Left            =   -74880
         TabIndex        =   26
         Top             =   360
         Width           =   8895
         Begin VB.CommandButton Cmd_ActualiserCirco 
            Default         =   -1  'True
            Height          =   375
            Left            =   6960
            Picture         =   "Frm_CommCircoCant.frx":047A
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Visualiser les circonscriptions associées canton sélectionné"
            Top             =   150
            Width           =   375
         End
         Begin VB.CommandButton Cmd_ValiderCommune 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   8400
            Picture         =   "Frm_CommCircoCant.frx":057C
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Enregistrer les modifications"
            Top             =   150
            Width           =   375
         End
         Begin VB.CommandButton Cmd_AnnulerCommune 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   7680
            Picture         =   "Frm_CommCircoCant.frx":067E
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Annuler les modifications"
            Top             =   150
            Width           =   375
         End
      End
      Begin VB.Frame Fra_OutilsCirco 
         Height          =   600
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   8895
         Begin VB.CommandButton cmd_ActualiserCanton 
            Height          =   375
            Left            =   6960
            Picture         =   "Frm_CommCircoCant.frx":0780
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Visualiser les cantons associés à la circonscription sélectionnée"
            Top             =   150
            Width           =   375
         End
         Begin VB.CommandButton Cmd_AnnulerCirco 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   7680
            Picture         =   "Frm_CommCircoCant.frx":0882
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Annuler les modifications"
            Top             =   150
            Width           =   375
         End
         Begin VB.CommandButton Cmd_ValiderCirco 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   8400
            Picture         =   "Frm_CommCircoCant.frx":0984
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Enregistrer les modifications"
            Top             =   150
            Width           =   375
         End
      End
      Begin VB.Frame Fra_Canton 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   8895
         Begin VB.Data Dat_Circo 
            Caption         =   "Dat_Circo"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   840
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   2160
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Data Dat_CircoComm 
            Caption         =   "Dat_CircoComm"
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
            Top             =   2040
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.PictureBox GAUCHE 
            Height          =   495
            Left            =   4080
            Picture         =   "Frm_CommCircoCant.frx":0A86
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   2040
            Width           =   615
         End
         Begin VB.PictureBox DROITE 
            Height          =   495
            Left            =   4080
            Picture         =   "Frm_CommCircoCant.frx":0EC8
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1080
            Width           =   615
         End
         Begin VB.ListBox Lst_Circo 
            Height          =   2400
            ItemData        =   "Frm_CommCircoCant.frx":130A
            Left            =   600
            List            =   "Frm_CommCircoCant.frx":130C
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   600
            Width           =   2895
         End
         Begin VB.ListBox Lst_CircoCommune 
            Height          =   2400
            ItemData        =   "Frm_CommCircoCant.frx":130E
            Left            =   5280
            List            =   "Frm_CommCircoCant.frx":1310
            Sorted          =   -1  'True
            TabIndex        =   19
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Liste des circonscriptions"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   960
            TabIndex        =   21
            Top             =   225
            Width           =   2385
         End
         Begin VB.Label Label2 
            Caption         =   "Liste des circonscritions faisant partie de la commune"
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
            Left            =   5520
            TabIndex        =   20
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.Frame Fra_Commune 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   -74880
         TabIndex        =   8
         Top             =   1080
         Width           =   8895
         Begin VB.PictureBox DROITE_Comm 
            Height          =   495
            Left            =   4080
            Picture         =   "Frm_CommCircoCant.frx":1312
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1080
            Width           =   615
         End
         Begin VB.PictureBox GAUCHE_Comm 
            Height          =   495
            Left            =   4080
            Picture         =   "Frm_CommCircoCant.frx":1754
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   2040
            Width           =   615
         End
         Begin VB.Data Dat_Cant 
            Caption         =   "Dat_Cant"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   840
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   2400
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Data Dat_CantComm 
            Caption         =   "Dat_CantComm"
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
            Top             =   2160
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.ListBox Lst_Cantons 
            Height          =   2400
            ItemData        =   "Frm_CommCircoCant.frx":1B96
            Left            =   600
            List            =   "Frm_CommCircoCant.frx":1B98
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   600
            Width           =   2895
         End
         Begin VB.ListBox Lst_CantComm 
            Height          =   2400
            ItemData        =   "Frm_CommCircoCant.frx":1B9A
            Left            =   5280
            List            =   "Frm_CommCircoCant.frx":1B9C
            TabIndex        =   12
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label Label3 
            Caption         =   "Liste des cantons faisant partie de la circonscription"
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
            Left            =   5640
            TabIndex        =   14
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label Label4 
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
            Height          =   165
            Left            =   1200
            TabIndex        =   13
            Top             =   225
            Width           =   1905
         End
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.Data Dat_Code 
         Caption         =   "Dat_Code"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmd_Quitter 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8640
         Picture         =   "Frm_CommCircoCant.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Label Label6 
      Caption         =   " Canton"
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
      TabIndex        =   33
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Circoncription"
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
      Left            =   960
      TabIndex        =   32
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lbl_num 
      Caption         =   "N° Commune"
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
      Left            =   960
      TabIndex        =   31
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lbl_nom 
      Caption         =   "Commune"
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
      TabIndex        =   30
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "Frm_CommCircoCant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ActualiserCanton_Click()

Dim vntTemp As Variant
Dim requete As String
Dim code_circo As Integer

If Lst_CircoCommune.ListIndex > -1 Then
    Txt_Circo.Text = Lst_CircoCommune.Text
    Txt_Canton.Text = ""
    
    requete = "SELECT CIRC_NUM FROM CIRC WHERE CIRC_NOM LIKE """ & Lst_CircoCommune.Text & """"
    Dat_Code.RecordSource = requete
    Dat_Code.Refresh
    code_circo = Dat_Code.recordset(0)
    
    ' liste des cantons qui ne sont pas liés à la circonscription à modifier
    requete = "SELECT CANT_NOM FROM CANT WHERE (CANT_NUM NOT IN"
    requete = requete + " (SELECT CANT_NUM FROM APPA"
    requete = requete + " WHERE COMM_INSEE=" & Txt_Code_commune & " AND CIRC_NUM=" & code_circo & "))"
    Dat_Cant.RecordSource = requete
    Dat_Cant.Refresh
    
    ' vide le listbox
    Do While Lst_Cantons.ListCount <> 0
     Lst_Cantons.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_Cant.recordset.EOF
     vntTemp = Dat_Cant.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_Cantons.AddItem CStr(vntTemp)
     Dat_Cant.recordset.MoveNext
    Loop
    
    ' liste des cantons qui sont liés à la circonscription et à la commune à modifier
    requete = "SELECT DISTINCT(CANT_NOM) FROM CANT,APPA WHERE "
    requete = requete + "APPA.CANT_NUM=CANT.CANT_NUM AND "
    requete = requete + "APPA.CIRC_NUM=" & code_circo
    requete = requete + " AND APPA.COMM_INSEE=" & Txt_Code_commune
    Dat_CantComm.RecordSource = requete
    Dat_CantComm.Refresh
    
    ' vide le listbox
    Do While Lst_CantComm.ListCount <> 0
     Lst_CantComm.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_CantComm.recordset.EOF
     vntTemp = Dat_CantComm.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_CantComm.AddItem CStr(vntTemp)
     Dat_CantComm.recordset.MoveNext
    Loop
End If

End Sub


Private Sub Cmd_ActualiserCirco_Click()
Dim vntTemp As Variant
Dim requete As String
Dim code_canton As Integer

If Lst_CantComm.ListIndex > -1 Then
    Txt_Canton.Text = Lst_CantComm.Text
    Txt_Circo.Text = ""
    
    requete = "SELECT CANT_NUM FROM CANT WHERE CANT_NOM LIKE """ & Lst_CantComm.Text & """"
    Dat_Code.RecordSource = requete
    Dat_Code.Refresh
    code_canton = Dat_Code.recordset(0)
    
    ' liste des circos qui ne sont pas liés au canton à modifier
    requete = "SELECT CIRC_NOM FROM CIRC WHERE (CIRC_NUM NOT IN"
    requete = requete + " (SELECT CIRC_NUM FROM APPA"
    requete = requete + " WHERE COMM_INSEE=" & Txt_Code_commune & " AND CANT_NUM=" & code_canton & "))"
    Dat_Circo.RecordSource = requete
    Dat_Circo.Refresh
    
    ' vide le listbox
    Do While Lst_Circo.ListCount <> 0
     Lst_Circo.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_Circo.recordset.EOF
     vntTemp = Dat_Circo.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_Circo.AddItem CStr(vntTemp)
     Dat_Circo.recordset.MoveNext
    Loop
    
    ' liste des circos qui sont liés au canton et à la commune à modifier
    requete = "SELECT DISTINCT(CIRC_NOM) FROM CIRC,APPA WHERE "
    requete = requete + "APPA.CIRC_NUM=CIRC.CIRC_NUM AND "
    requete = requete + "APPA.CANT_NUM=" & code_canton
    requete = requete + " AND APPA.COMM_INSEE=" & Txt_Code_commune
    Dat_CircoComm.RecordSource = requete
    Dat_CircoComm.Refresh
    
    ' vide le listbox
    Do While Lst_CircoCommune.ListCount <> 0
     Lst_CircoCommune.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_CircoComm.recordset.EOF
     vntTemp = Dat_CircoComm.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_CircoCommune.AddItem CStr(vntTemp)
     Dat_CircoComm.recordset.MoveNext
    Loop
End If

End Sub


Private Sub Cmd_AnnulerCirco_Click()
Dim vntTemp As Variant
Dim requete As String
    
    ' liste des cantons qui ne sont pas liés à la commune à modifier
    requete = "SELECT CANT_NOM FROM CANT WHERE (CANT_NUM NOT IN"
    requete = requete + " (SELECT CANT_NUM FROM APPA"
    requete = requete + " WHERE COMM_INSEE=" & Txt_Code_commune & "))"
    Dat_Cant.RecordSource = requete
    Dat_Cant.Refresh
    
    ' vide le listbox
    Do While Lst_Cantons.ListCount <> 0
     Lst_Cantons.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_Cant.recordset.EOF
     vntTemp = Dat_Cant.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_Cantons.AddItem CStr(vntTemp)
     Dat_Cant.recordset.MoveNext
    Loop
    
    ' liste des cantons qui sont liés à la commune à modifier
    requete = "SELECT DISTINCT(CANT_NOM) FROM CANT,APPA WHERE "
    requete = requete + "APPA.CANT_NUM=CANT.CANT_NUM AND "
    requete = requete + "APPA.COMM_INSEE=" & Txt_Code_commune
    Dat_CantComm.RecordSource = requete
    Dat_CantComm.Refresh
    
    ' vide le listbox
    Do While Lst_CantComm.ListCount <> 0
     Lst_CantComm.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_CantComm.recordset.EOF
     vntTemp = Dat_CantComm.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_CantComm.AddItem CStr(vntTemp)
     Dat_CantComm.recordset.MoveNext
    Loop

    Txt_Circo.Text = ""
    Txt_Canton.Text = ""
End Sub


Private Sub Cmd_AnnulerCommune_Click()
Dim vntTemp As Variant
Dim requete As String
    
    ' liste des circos qui ne sont pas liés à la commune à modifier
    requete = "SELECT CIRC_NOM FROM CIRC WHERE (CIRC_NUM NOT IN"
    requete = requete + " (SELECT CIRC_NUM FROM APPA"
    requete = requete + " WHERE COMM_INSEE=" & Txt_Code_commune & "))"
    Dat_Circo.RecordSource = requete
    Dat_Circo.Refresh
    
    ' vide le listbox
    Do While Lst_Circo.ListCount <> 0
     Lst_Circo.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_Circo.recordset.EOF
     vntTemp = Dat_Circo.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_Circo.AddItem CStr(vntTemp)
     Dat_Circo.recordset.MoveNext
    Loop
    
    ' liste des circos qui sont liés à la commune à modifier
    requete = "SELECT DISTINCT(CIRC_NOM) FROM CIRC,APPA WHERE "
    requete = requete + "APPA.CIRC_NUM=CIRC.CIRC_NUM AND "
    requete = requete + "APPA.COMM_INSEE=" & Txt_Code_commune
    Dat_CircoComm.RecordSource = requete
    Dat_CircoComm.Refresh
    
    ' vide le listbox
    Do While Lst_CircoCommune.ListCount <> 0
     Lst_CircoCommune.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_CircoComm.recordset.EOF
     vntTemp = Dat_CircoComm.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_CircoCommune.AddItem CStr(vntTemp)
     Dat_CircoComm.recordset.MoveNext
    Loop

    Txt_Circo.Text = ""
    Txt_Canton.Text = ""
End Sub


Private Sub cmd_Quitter_Click()
  Unload Me
End Sub


Private Sub DROITE_Click()

Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

    If Lst_Circo.ListIndex > -1 Then
        ind = Lst_Circo.ListIndex
        Lst_CircoCommune.AddItem Lst_Circo.List(ind)
        Lst_Circo.RemoveItem ind
    End If

End Sub

Private Sub DROITE_Comm_Click()
Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

    If Lst_Cantons.ListIndex > -1 Then
        ind = Lst_Cantons.ListIndex
        Lst_CantComm.AddItem Lst_Cantons.List(ind)
        Lst_Cantons.RemoveItem ind
    End If
End Sub

Private Sub Form_Load()

  Dat_Code.DatabaseName = CheminBase
  Dat_Circo.DatabaseName = CheminBase
  Dat_CircoComm.DatabaseName = CheminBase
  Dat_Cant.DatabaseName = CheminBase
  Dat_CantComm.DatabaseName = CheminBase
  
End Sub

Private Sub GAUCHE_Click()

Dim ind As Integer
  'Déplace un élément de la liste de droite vers la liste de gauche
  If Lst_CircoCommune.ListIndex > -1 Then
    ind = Lst_CircoCommune.ListIndex
    Lst_Circo.AddItem Lst_CircoCommune.List(ind)
    Lst_CircoCommune.RemoveItem ind
  End If
  
End Sub

Private Sub GAUCHE_Comm_Click()

Dim ind As Integer
  'Déplace un élément de la liste de droite vers la liste de gauche
  If Lst_CantComm.ListIndex > -1 Then
    ind = Lst_CantComm.ListIndex
    Lst_Cantons.AddItem Lst_CantComm.List(ind)
    Lst_CantComm.RemoveItem ind
  End If

End Sub

Private Sub Lst_CantComm_DblClick()

Dim ind As Integer
  'Déplace un élément de la liste de droite vers la liste de gauche
  If Lst_CantComm.ListIndex > -1 Then
    ind = Lst_CantComm.ListIndex
    Lst_Cantons.AddItem Lst_CantComm.List(ind)
    Lst_CantComm.RemoveItem ind
  End If
  
End Sub


Private Sub Lst_Cantons_DblClick()

Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

    If Lst_Cantons.ListIndex > -1 Then
        ind = Lst_Cantons.ListIndex
        Lst_CantComm.AddItem Lst_Cantons.List(ind)
        Lst_Cantons.RemoveItem ind
    End If

End Sub


Private Sub Lst_Circo_DblClick()

Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

    If Lst_Circo.ListIndex > -1 Then
        ind = Lst_Circo.ListIndex
        Lst_CircoCommune.AddItem Lst_Circo.List(ind)
        Lst_Circo.RemoveItem ind
    End If
    
End Sub


Private Sub Lst_CircoCommune_DblClick()

Dim ind As Integer
  'Déplace un élément de la liste de droite vers la liste de gauche
  If Lst_CircoCommune.ListIndex > -1 Then
    ind = Lst_CircoCommune.ListIndex
    Lst_Circo.AddItem Lst_CircoCommune.List(ind)
    Lst_CircoCommune.RemoveItem ind
  End If

End Sub



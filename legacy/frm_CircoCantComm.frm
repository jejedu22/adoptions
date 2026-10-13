VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_CircoCantComm 
   Caption         =   "Association d'une circonscription aux cantons et communes"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   Icon            =   "frm_CircoCantComm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Dat_Code 
      Caption         =   "Dat_Code"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   165
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Txt_Commune 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      TabIndex        =   31
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Txt_Canton 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   30
      Top             =   1200
      Width           =   2055
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
      TabCaption(0)   =   "Canton"
      TabPicture(0)   =   "frm_CircoCantComm.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_Canton"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra_OutilsCanton"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Commune"
      TabPicture(1)   =   "frm_CircoCantComm.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_OutilsCommune"
      Tab(1).Control(1)=   "Fra_Commune"
      Tab(1).ControlCount=   2
      Begin VB.Frame Fra_OutilsCommune 
         Height          =   600
         Left            =   -74880
         TabIndex        =   25
         Top             =   360
         Width           =   8895
         Begin VB.CommandButton Cmd_ActualiserCommune 
            Default         =   -1  'True
            Height          =   375
            Left            =   6960
            Picture         =   "frm_CircoCantComm.frx":047A
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Visualiser les cantons associés à la commune sélectionnée"
            Top             =   150
            Width           =   375
         End
         Begin VB.CommandButton Cmd_ValiderCommune 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   8400
            Picture         =   "frm_CircoCantComm.frx":057C
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Enregistrer les modifications"
            Top             =   150
            Width           =   375
         End
         Begin VB.CommandButton Cmd_AnnulerCommune 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   7680
            Picture         =   "frm_CircoCantComm.frx":067E
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Annuler les modifications"
            Top             =   150
            Width           =   375
         End
      End
      Begin VB.Frame Fra_OutilsCanton 
         Height          =   600
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   8895
         Begin VB.CommandButton cmd_ActualiserCanton 
            Height          =   375
            Left            =   6960
            Picture         =   "frm_CircoCantComm.frx":0780
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Visualiser les communes associées au canton sélectionné"
            Top             =   150
            Width           =   375
         End
         Begin VB.CommandButton Cmd_AnnulerCanton 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   7680
            Picture         =   "frm_CircoCantComm.frx":0882
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Annuler les modifications"
            Top             =   150
            Width           =   375
         End
         Begin VB.CommandButton Cmd_ValiderCanton 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   8400
            Picture         =   "frm_CircoCantComm.frx":0984
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
         Begin VB.Data Dat_Canton 
            Caption         =   "Dat_Canton"
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
            Width           =   2295
         End
         Begin VB.Data Dat_CantonCirco 
            Caption         =   "Dat_CantonCirco"
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
            Picture         =   "frm_CircoCantComm.frx":0A86
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
            Picture         =   "frm_CircoCantComm.frx":0EC8
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1080
            Width           =   615
         End
         Begin VB.ListBox Lst_Canton 
            Height          =   2400
            ItemData        =   "frm_CircoCantComm.frx":130A
            Left            =   600
            List            =   "frm_CircoCantComm.frx":130C
            TabIndex        =   19
            Top             =   600
            Width           =   2895
         End
         Begin VB.ListBox Lst_CantonCirco 
            Height          =   2400
            ItemData        =   "frm_CircoCantComm.frx":130E
            Left            =   5280
            List            =   "frm_CircoCantComm.frx":1310
            TabIndex        =   18
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label Label1 
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
            TabIndex        =   21
            Top             =   225
            Width           =   1680
         End
         Begin VB.Label Label2 
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
            TabIndex        =   20
            Top             =   120
            Width           =   2415
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
            Picture         =   "frm_CircoCantComm.frx":1312
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1080
            Width           =   615
         End
         Begin VB.PictureBox GAUCHE_Comm 
            Height          =   495
            Left            =   4080
            Picture         =   "frm_CircoCantComm.frx":1754
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2040
            Width           =   615
         End
         Begin VB.Data Dat_Commune 
            Caption         =   "Dat_Commune"
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
            Width           =   2415
         End
         Begin VB.Data Dat_CommuneCirco 
            Caption         =   "Dat_CommuneCirco"
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
            Width           =   2775
         End
         Begin VB.ListBox Lst_Commune 
            Height          =   2400
            ItemData        =   "frm_CircoCantComm.frx":1B96
            Left            =   600
            List            =   "frm_CircoCantComm.frx":1B98
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   600
            Width           =   2895
         End
         Begin VB.ListBox Lst_CommuneCirco 
            Height          =   2400
            ItemData        =   "frm_CircoCantComm.frx":1B9A
            Left            =   5280
            List            =   "frm_CircoCantComm.frx":1B9C
            TabIndex        =   10
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label Label3 
            Caption         =   "Liste des communes faisant partie de la circonscription"
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
            Height          =   165
            Left            =   1200
            TabIndex        =   13
            Top             =   225
            Width           =   1905
         End
      End
   End
   Begin VB.TextBox Txt_Code_circo 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Txt_Nom_circo 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   2535
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.CommandButton cmd_Quitter 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8640
         Picture         =   "frm_CircoCantComm.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_CircoCantComm 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6270
      Width           =   9390
      _ExtentX        =   16563
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
            TextSave        =   "13:08"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9887
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   " Commune"
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
      Left            =   4200
      TabIndex        =   29
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label5 
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
      Left            =   1080
      TabIndex        =   28
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl_num 
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
      Left            =   1080
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lbl_nom 
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
      Left            =   3840
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frm_CircoCantComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmd_ActualiserCanton_Click()

Dim vntTemp As Variant
Dim requete As String
Dim code_canton As Integer

If Lst_CantonCirco.ListIndex > -1 Then
    Txt_Canton.Text = Lst_CantonCirco.Text
    Txt_Commune.Text = ""
    
    requete = "SELECT CANT_NUM FROM CANT WHERE CANT_NOM LIKE """ & Lst_CantonCirco.Text & """"
    Dat_Code.RecordSource = requete
    Dat_Code.Refresh
    code_canton = Dat_Code.recordset(0)
    
    ' liste des communes qui ne sont pas liés à la circonscription à modifier
    requete = "SELECT COMM_NOM FROM COMM WHERE (COMM_INSEE NOT IN"
    requete = requete + " (SELECT COMM_INSEE FROM APPA"
    requete = requete + " WHERE CIRC_NUM=" & Txt_Code_circo & " AND CANT_NUM=" & code_canton & "))"
    Dat_Commune.RecordSource = requete
    Dat_Commune.Refresh
    
    ' vide le listbox
    Do While Lst_Commune.ListCount <> 0
     Lst_Commune.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_Commune.recordset.EOF
     vntTemp = Dat_Commune.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_Commune.AddItem CStr(vntTemp)
     Dat_Commune.recordset.MoveNext
    Loop
    
    ' liste des communes qui sont liés à la circonscription et au canton à modifier
    requete = "SELECT DISTINCT(COMM_NOM) FROM COMM,APPA WHERE "
    requete = requete + "APPA.COMM_INSEE=COMM.COMM_INSEE AND "
    requete = requete + "APPA.CIRC_NUM=" & Txt_Code_circo
    requete = requete + " AND APPA.CANT_NUM=" & code_canton
    Dat_CommuneCirco.RecordSource = requete
    Dat_CommuneCirco.Refresh
    
    ' vide le listbox
    Do While Lst_CommuneCirco.ListCount <> 0
     Lst_CommuneCirco.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_CommuneCirco.recordset.EOF
     vntTemp = Dat_CommuneCirco.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_CommuneCirco.AddItem CStr(vntTemp)
     Dat_CommuneCirco.recordset.MoveNext
    Loop
End If

End Sub

Private Sub Cmd_ActualiserCommune_Click()

Dim vntTemp As Variant
Dim requete As String
Dim code_commune As Integer

If Lst_CommuneCirco.ListIndex > -1 Then
    Txt_Commune.Text = Lst_CommuneCirco.Text
    Txt_Canton.Text = ""

    requete = "SELECT COMM_INSEE FROM COMM WHERE COMM_NOM LIKE """ & Lst_CommuneCirco.Text & """"
    Dat_Code.RecordSource = requete
    Dat_Code.Refresh
    code_commune = Dat_Code.recordset(0)
    
    ' liste des cantons qui ne sont pas liés à la circonscription à modifier
    requete = "SELECT CANT_NOM FROM CANT WHERE (CANT_NUM NOT IN"
    requete = requete + " (SELECT CANT_NUM FROM APPA"
    requete = requete + " WHERE CIRC_NUM=" & Txt_Code_circo & " AND COMM_INSEE=" & code_commune & "))"
    Dat_Canton.RecordSource = requete
    Dat_Canton.Refresh
    
    ' vide le listbox
    Do While Lst_Canton.ListCount <> 0
     Lst_Canton.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_Canton.recordset.EOF
     vntTemp = Dat_Canton.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_Canton.AddItem CStr(vntTemp)
     Dat_Canton.recordset.MoveNext
    Loop
    
    ' liste des cantons qui sont liés à la circonscription et à la commune à modifier
    requete = "SELECT DISTINCT(CANT_NOM) FROM CANT,APPA WHERE "
    requete = requete + "APPA.CANT_NUM=CANT.CANT_NUM AND "
    requete = requete + "APPA.CIRC_NUM=" & Txt_Code_circo
    requete = requete + " AND APPA.COMM_INSEE=" & code_commune
    Dat_CantonCirco.RecordSource = requete
    Dat_CantonCirco.Refresh
    
    ' vide le listbox
    Do While Lst_CantonCirco.ListCount <> 0
     Lst_CantonCirco.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_CantonCirco.recordset.EOF
     vntTemp = Dat_CantonCirco.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_CantonCirco.AddItem CStr(vntTemp)
     Dat_CantonCirco.recordset.MoveNext
    Loop
End If

End Sub

Private Sub Cmd_AnnulerCanton_Click()

Dim vntTemp As Variant
Dim requete As String
    
    ' liste des cantons qui ne sont pas liés à la circonscription à modifier
    requete = "SELECT CANT_NOM FROM CANT WHERE (CANT_NUM NOT IN"
    requete = requete + " (SELECT CANT_NUM FROM APPA"
    requete = requete + " WHERE CIRC_NUM=" & Txt_Code_circo & "))"
    Dat_Canton.RecordSource = requete
    Dat_Canton.Refresh
    
    ' vide le listbox
    Do While Lst_Canton.ListCount <> 0
     Lst_Canton.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_Canton.recordset.EOF
     vntTemp = Dat_Canton.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_Canton.AddItem CStr(vntTemp)
     Dat_Canton.recordset.MoveNext
    Loop
    
    ' liste des cantons qui sont liés à la circonscription et à la commune à modifier
    requete = "SELECT DISTINCT(CANT_NOM) FROM CANT,APPA WHERE "
    requete = requete + "APPA.CANT_NUM=CANT.CANT_NUM AND "
    requete = requete + "APPA.CIRC_NUM=" & Txt_Code_circo
    Dat_CantonCirco.RecordSource = requete
    Dat_CantonCirco.Refresh
    
    ' vide le listbox
    Do While Lst_CantonCirco.ListCount <> 0
     Lst_CantonCirco.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_CantonCirco.recordset.EOF
     vntTemp = Dat_CantonCirco.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_CantonCirco.AddItem CStr(vntTemp)
     Dat_CantonCirco.recordset.MoveNext
    Loop

    Txt_Commune.Text = ""
    Txt_Canton.Text = ""
    
End Sub

Private Sub Cmd_AnnulerCommune_Click()
    
Dim vntTemp As Variant
Dim requete As String
    
    ' liste des communes qui ne sont pas liés à la circonscription à modifier
    requete = "SELECT COMM_NOM FROM COMM WHERE (COMM_INSEE NOT IN"
    requete = requete + " (SELECT COMM_INSEE FROM APPA"
    requete = requete + " WHERE CIRC_NUM=" & Txt_Code_circo & "))"
    Dat_Commune.RecordSource = requete
    Dat_Commune.Refresh
    
    ' vide le listbox
    Do While Lst_Commune.ListCount <> 0
     Lst_Commune.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_Commune.recordset.EOF
     vntTemp = Dat_Commune.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_Commune.AddItem CStr(vntTemp)
     Dat_Commune.recordset.MoveNext
    Loop
    
    ' liste des communes qui sont liés à la circonscription et au canton à modifier
    requete = "SELECT DISTINCT(COMM_NOM) FROM COMM,APPA WHERE "
    requete = requete + "APPA.COMM_INSEE=COMM.COMM_INSEE AND "
    requete = requete + "APPA.CIRC_NUM=" & Txt_Code_circo
    Dat_CommuneCirco.RecordSource = requete
    Dat_CommuneCirco.Refresh
    
    ' vide le listbox
    Do While Lst_CommuneCirco.ListCount <> 0
     Lst_CommuneCirco.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_CommuneCirco.recordset.EOF
     vntTemp = Dat_CommuneCirco.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_CommuneCirco.AddItem CStr(vntTemp)
     Dat_CommuneCirco.recordset.MoveNext
    Loop

    Txt_Canton.Text = ""
    Txt_Commune.Text = ""
    
End Sub







Private Sub cmd_Quitter_Click()
    Unload Me
End Sub



Private Sub Cmd_ValiderCanton_Click()

Dim reponse, requete, selection As String
Dim i, ind, code_commune, code_canton As Integer

On Error GoTo Rupture_Integrite

If Txt_Commune.Text <> "" Then
  'Enregistre dans la base de données (table appartient) les relations entre la ville et son(ses) canton(s)
  reponse = MsgBox("Etes-vous sur de vouloir enregistrer les modifications ?", vbYesNo + vbQuestion + vbApplicationModal)
  If reponse = vbYes Then
    requete = "SELECT COMM_INSEE FROM COMM WHERE COMM_NOM='" & Lst_CommuneCirco.Text & "'"
    Dat_Code.RecordSource = requete
    Dat_Code.Refresh
    code_commune = Dat_Code.recordset(0)

    requete = "DELETE * FROM APPA WHERE CIRC_NUM=" & Txt_Code_circo
    requete = requete + " AND COMM_INSEE=" & code_commune
    Workspaces(0).Databases(0).Execute requete, dbFailOnError
    
    For i = 0 To Lst_CantonCirco.ListCount - 1
      selection = Lst_CantonCirco.List(i)
      requete = "SELECT CANT_NUM FROM CANT WHERE CANT_NOM='" & selection & "'"
      Dat_Code.RecordSource = requete
      Dat_Code.Refresh
      code_canton = Dat_Code.recordset(0)
      
      requete = "INSERT INTO APPA VALUES (" & code_canton & "," & code_commune & "," & Txt_Code_circo & ")"
      Workspaces(0).Databases(0).Execute requete, dbFailOnError
    Next i
    
    Cmd_AnnulerCommune_Click
    
  End If
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

Private Sub Cmd_ValiderCommune_Click()

Dim reponse, requete, selection As String
Dim i, code_commune, code_canton As Integer

On Error GoTo Rupture_Integrite

If Txt_Canton.Text <> "" Then
  'Enregistre dans la base de données (table appartient) les relations entre le canton et sa(ses) commune(s)
  reponse = MsgBox("Etes-vous sur de vouloir enregistrer les modifications ?", vbYesNo + vbQuestion + vbApplicationModal)
  If reponse = vbYes Then
    requete = "SELECT CANT_NUM FROM CANT WHERE CANT_NOM='" & Lst_CantonCirco.Text & "'"
    Dat_Code.RecordSource = requete
    Dat_Code.Refresh
    code_canton = Dat_Code.recordset(0)

    requete = "DELETE * FROM APPA WHERE CIRC_NUM=" & Txt_Code_circo
    requete = requete + " AND CANT_NUM=" & code_canton
    Workspaces(0).Databases(0).Execute requete, dbFailOnError
    
    For i = 0 To Lst_CommuneCirco.ListCount - 1
      selection = Lst_CommuneCirco.List(i)
      requete = "SELECT COMM_INSEE FROM COMM WHERE COMM_NOM='" & selection & "'"
      Dat_Code.RecordSource = requete
      Dat_Code.Refresh
      code_commune = Dat_Code.recordset(0)
      
      requete = "INSERT INTO APPA VALUES (" & code_canton & "," & code_commune & "," & Txt_Code_circo & ")"
      Workspaces(0).Databases(0).Execute requete, dbFailOnError
    Next i
    
    Cmd_AnnulerCanton_Click
  End If
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

    If Lst_Canton.ListIndex > -1 Then
        ind = Lst_Canton.ListIndex
        Lst_CantonCirco.AddItem Lst_Canton.List(ind)
        Lst_Canton.RemoveItem ind
    End If
    

End Sub



Private Sub DROITE_Comm_Click()

Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

    If Lst_Commune.ListIndex > -1 Then
        ind = Lst_Commune.ListIndex
        Lst_CommuneCirco.AddItem Lst_Commune.List(ind)
        Lst_Commune.RemoveItem ind
    End If

End Sub

Private Sub Form_Load()
    
    Dat_Code.DatabaseName = CheminBase
    Dat_Canton.DatabaseName = CheminBase
    Dat_CantonCirco.DatabaseName = CheminBase
    Dat_Commune.DatabaseName = CheminBase
    Dat_CommuneCirco.DatabaseName = CheminBase
    
End Sub

Private Sub GAUCHE_Click()

Dim ind As Integer
    'Déplace un élément de la liste de droite vers la liste de gauche
    If Lst_CantonCirco.ListIndex > -1 Then
      ind = Lst_CantonCirco.ListIndex
      Lst_Canton.AddItem Lst_CantonCirco.List(ind)
      Lst_CantonCirco.RemoveItem ind
    End If
  
End Sub

Private Sub GAUCHE_Comm_Click()

Dim ind As Integer
    'Déplace un élément de la liste de droite vers la liste de gauche
    If Lst_CommuneCirco.ListIndex > -1 Then
      ind = Lst_CommuneCirco.ListIndex
      Lst_Commune.AddItem Lst_CommuneCirco.List(ind)
      Lst_CommuneCirco.RemoveItem ind
    End If

End Sub


Private Sub Lst_Canton_DblClick()

Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

    If Lst_Canton.ListIndex > -1 Then
        ind = Lst_Canton.ListIndex
        Lst_CantonCirco.AddItem Lst_Canton.List(ind)
        Lst_Canton.RemoveItem ind
    End If
    
End Sub


Private Sub Lst_CantonCirco_DblClick()

Dim ind As Integer
    'Déplace un élément de la liste de droite vers la liste de gauche
    If Lst_CantonCirco.ListIndex > -1 Then
      ind = Lst_CantonCirco.ListIndex
      Lst_Canton.AddItem Lst_CantonCirco.List(ind)
      Lst_CantonCirco.RemoveItem ind
    End If
    
End Sub








Private Sub Lst_Commune_DblClick()

Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

    If Lst_Commune.ListIndex > -1 Then
        ind = Lst_Commune.ListIndex
        Lst_CommuneCirco.AddItem Lst_Commune.List(ind)
        Lst_Commune.RemoveItem ind
    End If


End Sub


Private Sub Lst_CommuneCirco_DblClick()

Dim ind As Integer
    'Déplace un élément de la liste de droite vers la liste de gauche
    If Lst_CommuneCirco.ListIndex > -1 Then
      ind = Lst_CommuneCirco.ListIndex
      Lst_Commune.AddItem Lst_CommuneCirco.List(ind)
      Lst_CommuneCirco.RemoveItem ind
    End If

End Sub



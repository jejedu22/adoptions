VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_CantCircoComm 
   Caption         =   "Association d'un canton aux circonscriptions et communes"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   Icon            =   "Frm_CantCircoComm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txt_Nom_canton 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Txt_Code_canton 
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
   Begin VB.TextBox Txt_Circo 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Txt_Commune 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
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
         Left            =   1800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   165
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmd_Quitter 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8640
         Picture         =   "Frm_CantCircoComm.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_CantCircoComm 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
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
            TextSave        =   "13:08"
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
      TabIndex        =   11
      Top             =   1680
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Circonscription"
      TabPicture(0)   =   "Frm_CantCircoComm.frx":0544
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_OutilsCirco"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra_Canton"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Commune"
      TabPicture(1)   =   "Frm_CantCircoComm.frx":0560
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_OutilsCommune"
      Tab(1).Control(1)=   "Fra_Commune"
      Tab(1).ControlCount=   2
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
         TabIndex        =   27
         Top             =   1080
         Width           =   8895
         Begin VB.Data Dat_CommuneCanton 
            Caption         =   "Dat_CommuneCanton"
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
         Begin VB.PictureBox GAUCHE_Comm 
            Height          =   495
            Left            =   4080
            Picture         =   "Frm_CantCircoComm.frx":057C
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   2040
            Width           =   615
         End
         Begin VB.PictureBox DROITE_Comm 
            Height          =   495
            Left            =   4080
            Picture         =   "Frm_CantCircoComm.frx":09BE
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1080
            Width           =   615
         End
         Begin VB.ListBox Lst_Commune 
            Height          =   2400
            ItemData        =   "Frm_CantCircoComm.frx":0E00
            Left            =   600
            List            =   "Frm_CantCircoComm.frx":0E02
            Sorted          =   -1  'True
            TabIndex        =   30
            Top             =   600
            Width           =   2895
         End
         Begin VB.ListBox Lst_CommuneCanton 
            Height          =   2400
            ItemData        =   "Frm_CantCircoComm.frx":0E04
            Left            =   5280
            List            =   "Frm_CantCircoComm.frx":0E06
            TabIndex        =   31
            Top             =   600
            Width           =   2895
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
            TabIndex        =   33
            Top             =   225
            Width           =   1905
         End
         Begin VB.Label Label3 
            Caption         =   "Liste des communes faisant partie du canton"
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
            TabIndex        =   32
            Top             =   120
            Width           =   2175
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
         TabIndex        =   20
         Top             =   1080
         Width           =   8895
         Begin VB.PictureBox DROITE 
            Height          =   495
            Left            =   4080
            Picture         =   "Frm_CantCircoComm.frx":0E08
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   1080
            Width           =   615
         End
         Begin VB.PictureBox GAUCHE 
            Height          =   495
            Left            =   4080
            Picture         =   "Frm_CantCircoComm.frx":124A
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   2040
            Width           =   615
         End
         Begin VB.Data Dat_CircoCanton 
            Caption         =   "Dat_CircoCanton"
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
         Begin VB.ListBox Lst_Circo 
            Height          =   2400
            ItemData        =   "Frm_CantCircoComm.frx":168C
            Left            =   600
            List            =   "Frm_CantCircoComm.frx":168E
            Sorted          =   -1  'True
            TabIndex        =   23
            Top             =   600
            Width           =   2895
         End
         Begin VB.ListBox Lst_CircoCanton 
            Height          =   2400
            ItemData        =   "Frm_CantCircoComm.frx":1690
            Left            =   5280
            List            =   "Frm_CantCircoComm.frx":1692
            Sorted          =   -1  'True
            TabIndex        =   24
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Liste des circonscritions faisant partie du canton"
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
            TabIndex        =   26
            Top             =   120
            Width           =   2415
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
            TabIndex        =   25
            Top             =   225
            Width           =   2385
         End
      End
      Begin VB.Frame Fra_OutilsCirco 
         Height          =   600
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   8895
         Begin VB.CommandButton Cmd_ValiderCirco 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   8400
            Picture         =   "Frm_CantCircoComm.frx":1694
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Enregistrer les modifications"
            Top             =   150
            Width           =   375
         End
         Begin VB.CommandButton Cmd_AnnulerCirco 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   7680
            Picture         =   "Frm_CantCircoComm.frx":1796
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Annuler les modifications"
            Top             =   150
            Width           =   375
         End
         Begin VB.CommandButton cmd_ActualiserCirco 
            Height          =   375
            Left            =   6960
            Picture         =   "Frm_CantCircoComm.frx":1898
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Visualiser les communes associées au canton sélectionné"
            Top             =   150
            Width           =   375
         End
      End
      Begin VB.Frame Fra_OutilsCommune 
         Height          =   600
         Left            =   -74880
         TabIndex        =   12
         Top             =   360
         Width           =   8895
         Begin VB.CommandButton Cmd_AnnulerCommune 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   7680
            Picture         =   "Frm_CantCircoComm.frx":199A
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Annuler les modifications"
            Top             =   150
            Width           =   375
         End
         Begin VB.CommandButton Cmd_ValiderCommune 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   8400
            Picture         =   "Frm_CantCircoComm.frx":1A9C
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Enregistrer les modifications"
            Top             =   150
            Width           =   375
         End
         Begin VB.CommandButton Cmd_ActualiserCommune 
            Default         =   -1  'True
            Height          =   375
            Left            =   6960
            Picture         =   "Frm_CantCircoComm.frx":1B9E
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Visualiser les cantons associés à la commune sélectionnée"
            Top             =   150
            Width           =   375
         End
      End
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
      Left            =   4845
      TabIndex        =   9
      Top             =   720
      Width           =   615
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
      Left            =   1080
      TabIndex        =   8
      Top             =   720
      Width           =   975
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
      Left            =   1080
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
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
      Left            =   4800
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "Frm_CantCircoComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ActualiserCirco_Click()

Dim vntTemp As Variant
Dim requete As String
Dim code_circo As Integer

If Lst_CircoCanton.ListIndex > -1 Then
    Txt_Circo.Text = Lst_CircoCanton.Text
    Txt_Commune.Text = ""
    
    requete = "SELECT CIRC_NUM FROM CIRC WHERE CIRC_NOM='" & Lst_CircoCanton.Text & "'"
    Dat_Code.RecordSource = requete
    Dat_Code.Refresh
    code_circo = Dat_Code.recordset(0)
    
    ' liste des communes qui ne sont pas liés au canton à modifier
    requete = "SELECT COMM_NOM FROM COMM WHERE (COMM_INSEE NOT IN"
    requete = requete + " (SELECT COMM_INSEE FROM APPA"
    requete = requete + " WHERE CIRC_NUM=" & code_circo & " AND CANT_NUM=" & Txt_Code_canton & "))"
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
    requete = requete + "APPA.CIRC_NUM=" & code_circo
    requete = requete + " AND APPA.CANT_NUM=" & Txt_Code_canton
    Dat_CommuneCanton.RecordSource = requete
    Dat_CommuneCanton.Refresh
    
    ' vide le listbox
    Do While Lst_CommuneCanton.ListCount <> 0
     Lst_CommuneCanton.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_CommuneCanton.recordset.EOF
     vntTemp = Dat_CommuneCanton.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_CommuneCanton.AddItem CStr(vntTemp)
     Dat_CommuneCanton.recordset.MoveNext
    Loop
End If

End Sub

Private Sub Cmd_ActualiserCommune_Click()

Dim vntTemp As Variant
Dim requete As String
Dim code_commune As Integer

If Lst_CommuneCanton.ListIndex > -1 Then
    Txt_Commune.Text = Lst_CommuneCanton.Text
    Txt_Circo.Text = ""

    requete = "SELECT COMM_INSEE FROM COMM WHERE COMM_NOM ='" & Lst_CommuneCanton.Text & "'"
    Dat_Code.RecordSource = requete
    Dat_Code.Refresh
    code_commune = Dat_Code.recordset(0)
    
    ' liste des circos qui ne sont pas liés au canton à modifier
    requete = "SELECT CIRC_NOM FROM CIRC WHERE (CIRC_NUM NOT IN"
    requete = requete + " (SELECT CIRC_NUM FROM APPA"
    requete = requete + " WHERE CANT_NUM=" & Txt_Code_canton & " AND COMM_INSEE=" & code_commune & "))"
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
    requete = requete + "APPA.CANT_NUM=" & Txt_Code_canton
    requete = requete + " AND APPA.COMM_INSEE=" & code_commune
    Dat_CircoCanton.RecordSource = requete
    Dat_CircoCanton.Refresh
    
    ' vide le listbox
    Do While Lst_CircoCanton.ListCount <> 0
     Lst_CircoCanton.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_CircoCanton.recordset.EOF
     vntTemp = Dat_CircoCanton.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_CircoCanton.AddItem CStr(vntTemp)
     Dat_CircoCanton.recordset.MoveNext
    Loop
End If

End Sub

Private Sub Cmd_AnnulerCirco_Click()

Dim vntTemp As Variant
Dim requete As String
    
    ' liste des circos qui ne sont pas liés àau canton à modifier
    requete = "SELECT CIRC_NOM FROM CIRC WHERE (CIRC_NUM NOT IN"
    requete = requete + " (SELECT CIRC_NUM FROM APPA"
    requete = requete + " WHERE CANT_NUM=" & Txt_Code_canton & "))"
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
    
    ' liste des circos qui sont liées au canton
    requete = "SELECT DISTINCT(CIRC_NOM) FROM CIRC,APPA WHERE "
    requete = requete + "APPA.CIRC_NUM=CIRC.CIRC_NUM AND "
    requete = requete + "APPA.CANT_NUM=" & Txt_Code_canton
    Dat_CircoCanton.RecordSource = requete
    Dat_CircoCanton.Refresh
    
    ' vide le listbox
    Do While Lst_CircoCanton.ListCount <> 0
     Lst_CircoCanton.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_CircoCanton.recordset.EOF
     vntTemp = Dat_CircoCanton.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_CircoCanton.AddItem CStr(vntTemp)
     Dat_CircoCanton.recordset.MoveNext
    Loop

    Txt_Commune.Text = ""
    Txt_Circo.Text = ""
End Sub

Private Sub Cmd_AnnulerCommune_Click()

Dim vntTemp As Variant
Dim requete As String
    
    ' liste des communes qui ne sont pas liés à la circonscription à modifier
    requete = "SELECT COMM_NOM FROM COMM WHERE (COMM_INSEE NOT IN"
    requete = requete + " (SELECT COMM_INSEE FROM APPA"
    requete = requete + " WHERE CANT_NUM=" & Txt_Code_canton & "))"
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
    requete = requete + "APPA.CANT_NUM=" & Txt_Code_canton
    Dat_CommuneCanton.RecordSource = requete
    Dat_CommuneCanton.Refresh
    
    ' vide le listbox
    Do While Lst_CommuneCanton.ListCount <> 0
     Lst_CommuneCanton.RemoveItem (0)
    Loop
    
    ' remplissage du listbox avec les valeurs de la requète
    Do While Not Dat_CommuneCanton.recordset.EOF
     vntTemp = Dat_CommuneCanton.recordset(0)
     If IsNull(vntTemp) Then vntTemp = ""
     Lst_CommuneCanton.AddItem CStr(vntTemp)
     Dat_CommuneCanton.recordset.MoveNext
    Loop

    Txt_Circo.Text = ""
    Txt_Commune.Text = ""
End Sub


Private Sub cmd_Quitter_Click()
    Unload Me
End Sub

Private Sub Cmd_ValiderCirco_Click()

Dim reponse, requete, selection As String
Dim i, code_commune, code_circo As Integer

On Error GoTo Rupture_Integrite

If Txt_Commune.Text <> "" Then
  'Enregistre dans la base de données (table appartient) les relations entre la ville et son(ses) canton(s)
  reponse = MsgBox("Etes-vous sur de vouloir enregistrer les modifications ?", vbYesNo + vbQuestion + vbApplicationModal)
  If reponse = vbYes Then
    requete = "SELECT COMM_INSEE FROM COMM WHERE COMM_NOM='" & Lst_CommuneCanton.Text & "'"
    Dat_Code.RecordSource = requete
    Dat_Code.Refresh
    code_commune = Dat_Code.recordset(0)

    requete = "DELETE * FROM APPA WHERE CANT_NUM=" & Txt_Code_canton
    requete = requete + " AND COMM_INSEE=" & code_commune
    Workspaces(0).Databases(0).Execute requete, dbFailOnError
    
    For i = 0 To Lst_CircoCanton.ListCount - 1
      selection = Lst_CircoCanton.List(i)
      requete = "SELECT CIRC_NUM FROM CIRC WHERE CIRC_NOM='" & selection & "'"
      Dat_Code.RecordSource = requete
      Dat_Code.Refresh
      code_circo = Dat_Code.recordset(0)
      
      requete = "INSERT INTO APPA VALUES (" & Txt_Code_canton & "," & code_commune & "," & code_circo & ")"
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
Dim i, code_commune, code_circo As Integer

On Error GoTo Rupture_Integrite

If Txt_Circo.Text <> "" Then
  'Enregistre dans la base de données (table appartient) les relations entre le canton et sa(ses) commune(s)
  reponse = MsgBox("Etes-vous sur de vouloir enregistrer les modifications ?", vbYesNo + vbQuestion + vbApplicationModal)
  If reponse = vbYes Then
    requete = "SELECT CIRC_NUM FROM CIRC WHERE CIRC_NOM='" & Lst_CircoCanton.Text & "'"
    Dat_Code.RecordSource = requete
    Dat_Code.Refresh
    code_circo = Dat_Code.recordset(0)

    requete = "DELETE * FROM APPA WHERE CIRC_NUM=" & code_circo
    requete = requete + " AND CANT_NUM=" & Txt_Code_canton
    Workspaces(0).Databases(0).Execute requete, dbFailOnError
    
    For i = 0 To Lst_CommuneCanton.ListCount - 1
      selection = Lst_CommuneCanton.List(i)
      requete = "SELECT COMM_INSEE FROM COMM WHERE COMM_NOM='" & selection & "'"
      Dat_Code.RecordSource = requete
      Dat_Code.Refresh
      code_commune = Dat_Code.recordset(0)
      
      requete = "INSERT INTO APPA VALUES (" & Txt_Code_canton & "," & code_commune & "," & code_circo & ")"
      Workspaces(0).Databases(0).Execute requete, dbFailOnError
    Next i
    
    Cmd_AnnulerCirco_Click
    
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

    If Lst_Circo.ListIndex > -1 Then
        ind = Lst_Circo.ListIndex
        Lst_CircoCanton.AddItem Lst_Circo.List(ind)
        Lst_Circo.RemoveItem ind
    End If

End Sub

Private Sub DROITE_Comm_Click()

Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

    If Lst_Commune.ListIndex > -1 Then
        ind = Lst_Commune.ListIndex
        Lst_CommuneCanton.AddItem Lst_Commune.List(ind)
        Lst_Commune.RemoveItem ind
    End If
    
End Sub

Private Sub Form_Load()

    Dat_Code.DatabaseName = CheminBase
    Dat_Circo.DatabaseName = CheminBase
    Dat_CircoCanton.DatabaseName = CheminBase
    Dat_Commune.DatabaseName = CheminBase
    Dat_CommuneCanton.DatabaseName = CheminBase
        
End Sub

Private Sub GAUCHE_Click()

Dim ind As Integer
    'Déplace un élément de la liste de droite vers la liste de gauche
    If Lst_CircoCanton.ListIndex > -1 Then
      ind = Lst_CircoCanton.ListIndex
      Lst_Circo.AddItem Lst_CircoCanton.List(ind)
      Lst_CircoCanton.RemoveItem ind
    End If
  
End Sub

Private Sub GAUCHE_Comm_Click()
Dim ind As Integer
    'Déplace un élément de la liste de droite vers la liste de gauche
    If Lst_CommuneCanton.ListIndex > -1 Then
      ind = Lst_CommuneCanton.ListIndex
      Lst_Commune.AddItem Lst_CommuneCanton.List(ind)
      Lst_CommuneCanton.RemoveItem ind
    End If

End Sub

Private Sub Lst_Circo_DblClick()

Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

    If Lst_Circo.ListIndex > -1 Then
        ind = Lst_Circo.ListIndex
        Lst_CircoCanton.AddItem Lst_Circo.List(ind)
        Lst_Circo.RemoveItem ind
    End If

End Sub


Private Sub Lst_CircoCanton_DblClick()

Dim ind As Integer
    'Déplace un élément de la liste de droite vers la liste de gauche
    If Lst_CircoCanton.ListIndex > -1 Then
      ind = Lst_CircoCanton.ListIndex
      Lst_Circo.AddItem Lst_CircoCanton.List(ind)
      Lst_CircoCanton.RemoveItem ind
    End If
    
End Sub


Private Sub Lst_Commune_DblClick()

Dim ind As Integer
  'Fait passer un élément de la liste de gauche dans celle de droite

    If Lst_Commune.ListIndex > -1 Then
        ind = Lst_Commune.ListIndex
        Lst_CommuneCanton.AddItem Lst_Commune.List(ind)
        Lst_Commune.RemoveItem ind
    End If
    
End Sub


Private Sub Lst_CommuneCanton_DblClick()

Dim ind As Integer
    'Déplace un élément de la liste de droite vers la liste de gauche
    If Lst_CommuneCanton.ListIndex > -1 Then
      ind = Lst_CommuneCanton.ListIndex
      Lst_Commune.AddItem Lst_CommuneCanton.List(ind)
      Lst_CommuneCanton.RemoveItem ind
    End If

End Sub



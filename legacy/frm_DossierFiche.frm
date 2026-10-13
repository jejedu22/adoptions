VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_DossierFiche 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dossier"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_DossierFiche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fra_dossiers 
      Caption         =   "Dossiers d'adoptions"
      Height          =   4335
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   9255
      Begin VB.Data Dat_ReqSituationFamiliale 
         Caption         =   "Dat_ReqSitF"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4035
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Data Dat_Circonscriptions 
         Caption         =   "Dat_Circonscriptions"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   6960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4035
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Data Dat_Commune 
         Caption         =   "Dat_Commune"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   4560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4035
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data Dat_Canton 
         Caption         =   "Dat_Canton"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4035
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Frame Fra_Situation 
         Caption         =   "Situation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   33
         Top             =   2880
         Width           =   9015
         Begin VB.TextBox Txt_regmat 
            Height          =   285
            Left            =   5400
            MaxLength       =   20
            TabIndex        =   12
            Top             =   240
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker DTP_DateMariage 
            Height          =   255
            Left            =   1440
            TabIndex        =   13
            ToolTipText     =   "Date de mariage"
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yyy"
            Format          =   58851331
            CurrentDate     =   36284
         End
         Begin VB.TextBox Txt_LieuMariage 
            Height          =   285
            Left            =   5400
            MaxLength       =   20
            TabIndex        =   14
            ToolTipText     =   "Lieu de mariage"
            Top             =   600
            Width           =   1935
         End
         Begin VB.ComboBox Cmb_SituationFamiliale 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   11
            ToolTipText     =   "Situation familiale"
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lbl_RegimeMatrimonial 
            Caption         =   "Régime Matrimonial"
            Height          =   255
            Left            =   3840
            TabIndex        =   37
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lbl_datemariage 
            Caption         =   "Date de mariage"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lbl_LieuMariage 
            Caption         =   "Lieu de mariage"
            Height          =   255
            Left            =   3840
            TabIndex        =   35
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lbl_SituaFamil 
            Caption         =   "Situation familiale"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.ComboBox Cmb_Commune 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Commune"
         Top             =   1560
         Width           =   2895
      End
      Begin VB.ComboBox Cmb_Canton 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Canton"
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox Txt_Ville 
         Height          =   285
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "Ville du foyer demandeur"
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Txt_CP 
         Height          =   285
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   3
         ToolTipText     =   "Code postal du foyer demandeur"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Txt_adr2 
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   2
         ToolTipText     =   "Adresse du foyer demandeur"
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox Txt_adr1 
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   1
         ToolTipText     =   "Adresse du foyer demandeur"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox Txt_num 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         ToolTipText     =   "Entrez le numéro pour ce dossier"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Frame Fra_Identification 
         Caption         =   "Identification"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   9015
         Begin VB.Label lbl_code 
            Caption         =   "Numéro dossier"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Fra_Caracteristiques 
         Caption         =   "Caractéristiques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   9015
         Begin VB.ComboBox Cmb_Circo 
            Height          =   315
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   10
            ToolTipText     =   "Canton"
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox Txt_Departement 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7680
            MaxLength       =   2
            TabIndex        =   7
            ToolTipText     =   "Département"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Txt_Tel 
            Height          =   285
            Left            =   1080
            MaxLength       =   14
            TabIndex        =   5
            ToolTipText     =   "Numéro de téléphone du foyer demandeur"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.ComboBox Cmb_reins 
            Height          =   315
            Left            =   5880
            Style           =   2  'Dropdown List
            TabIndex        =   6
            ToolTipText     =   "Indique si le dossier est une réinstruction ou non"
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lbl_CodeDepartement 
            Caption         =   "Département"
            Enabled         =   0   'False
            Height          =   255
            Left            =   6600
            TabIndex        =   32
            Top             =   240
            Width           =   1095
         End
         Begin VB.Line Lin_Separateur 
            BorderColor     =   &H80000016&
            X1              =   4800
            X2              =   4800
            Y1              =   120
            Y2              =   1800
         End
         Begin VB.Label lbl_Tel 
            Caption         =   "Téléphone"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lbl_adr1 
            Caption         =   "Adresse 1"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbl_reinstruction 
            Caption         =   "Réinstruction"
            Height          =   255
            Left            =   4920
            TabIndex        =   29
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lbl_adr2 
            Caption         =   "Adresse 2"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lbl_CP 
            Caption         =   "Code Postal"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lbl_ville 
            Caption         =   "Ville"
            Height          =   255
            Left            =   2160
            TabIndex        =   26
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lbl_Commune 
            Caption         =   "Commune"
            Height          =   255
            Left            =   4920
            TabIndex        =   25
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lbl_Canton 
            Caption         =   "Canton"
            Height          =   255
            Left            =   4920
            TabIndex        =   24
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lbl_Circonscription 
            Caption         =   "Circonscription"
            Height          =   255
            Left            =   4920
            TabIndex        =   23
            Top             =   1320
            Width           =   1215
         End
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Quitter 
         Height          =   375
         Left            =   7920
         Picture         =   "frm_DossierFiche.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   8640
         Picture         =   "frm_DossierFiche.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_DossiersFiche 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   4995
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "18/07/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "09:53"
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
End
Attribute VB_Name = "frm_DossierFiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private QuitterCroix As Boolean

Private Sub Cmb_Canton_Click()

    'With Dat_Circonscriptions
    '    .RecordSource = "SELECT CIRC.* FROM CIRC,APPA WHERE CIRC.CIRC_NUM=APPA.CIRC_NUM" + _
    '                  " AND APPA.CANT_NUM=" & CodeEntreCotes(Cmb_Canton) + _
    '                  " AND APPA.COMM_INSEE=" & CodeEntreCotes(Cmb_Commune)
    '    .Refresh
    'End With
    'Txt_circo.Text = Dat_Circonscriptions.Recordset("CIRC_NOM")


End Sub

Private Sub Cmb_Canton_Gotfocus()
    If (stb_DossiersFiche.Panels(4).Text = "CREATION") Or (stb_DossiersFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DossiersFiche.SimpleText = "Entrez le canton du foyer demandeur"
    End If
End Sub


Private Sub Cmb_Commune_Click()
    
    'met à jour la liste des cantons suivant la commune choisie
    With Dat_Canton
        .RecordSource = "SELECT ALL DISTINCT (CANT.CANT_NUM), CANT.CANT_NOM FROM CANT,COMM,CACO WHERE " + _
                              "CANT.CANT_NUM = CACO.CANT_NUM AND " + _
                              "CACO.COMM_INSEE=" & CodeEntreCotes(Cmb_Commune) + _
                              " ORDER BY CANT_NOM"
        .Refresh
        Cmb_Canton.Clear
        While Not Dat_Canton.Recordset.EOF
            Cmb_Canton.AddItem (.Recordset.Fields("CANT_NOM") + " <" + CStr(.Recordset.Fields("CANT_NUM")) + ">")
            .Recordset.MoveNext
        Wend
    End With
    Cmb_Canton.ListIndex = 0
    
    'met à jour la liste des circonscriptions suivant la commune choisie
    With Dat_Circonscriptions
        .RecordSource = "SELECT ALL DISTINCT (CIRC.CIRC_NUM), CIRC.CIRC_NOM FROM CIRC,COMM,CICO WHERE " + _
                              "CIRC.CIRC_NUM = CICO.CIRC_NUM AND " + _
                              "CICO.COMM_INSEE=" & CodeEntreCotes(Cmb_Commune) + _
                              " ORDER BY CIRC_NOM"
        .Refresh
        Cmb_Circo.Clear
        While Not Dat_Circonscriptions.Recordset.EOF
            Cmb_Circo.AddItem (.Recordset.Fields("CIRC_NOM") + " <" + CStr(.Recordset.Fields("CIRC_NUM")) + ">")
            .Recordset.MoveNext
        Wend
    End With
    Cmb_Circo.ListIndex = 0
    
End Sub

Private Sub Cmb_Commune_Gotfocus()
    If (stb_DossiersFiche.Panels(4).Text = "CREATION") Or (stb_DossiersFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DossiersFiche.SimpleText = "Entrez la commune du foyer demandeur"
    End If
End Sub

Private Sub Cmb_RegimeMatrimonial_Gotfocus()
    If (stb_DossiersFiche.Panels(4).Text = "CREATION") Or (stb_DossiersFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DossiersFiche.SimpleText = "Entrez le régime matrimonial du foyer demandeur (s'il y a lieu)"
    End If
End Sub

Private Sub Cmb_reins_Gotfocus()
    If (stb_DossiersFiche.Panels(4).Text = "CREATION") Or (stb_DossiersFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DossiersFiche.SimpleText = "S'agit-il d'une réinstruction : (O)ui ou (N)on"
    End If
End Sub

Private Sub Cmb_SituationFamiliale_Gotfocus()
    If (stb_DossiersFiche.Panels(4).Text = "CREATION") Or (stb_DossiersFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DossiersFiche.SimpleText = "Entrez la situation familiale du foyer demandeur"
    End If
End Sub



Private Sub cmd_Quitter_Click()
    QuitterCroix = False
    Unload Me
End Sub

Private Sub cmd_Valider_Click()
Const pkstrNomProcédure As String = "cmd_Valider_Click"
  Dim requete As String
  Dim sql As String
  Dim RsDema As Recordset
  Dim RsCoac As Recordset
  Dim RsFade As Recordset
  Dim NumEnf As String
  Dim NumDema As String
On Error GoTo Err_cmd_Valider_Click
        
  QuitterCroix = False
  
  Select Case stb_DossiersFiche.Panels(4).Text
    Case "CREATION"
      MousePointer = vbHourglass
      requete = "INSERT INTO DOSS VALUES("
      requete = requete + Iifch(Txt_num) + ","
      requete = requete + CodeEntreCotes(Cmb_SituationFamiliale.Text) + ","
      requete = requete + CodeEntreCotes(Cmb_Canton) + ","
      requete = requete + CodeEntreCotes(Cmb_Commune) + ","
      requete = requete + CodeEntreCotes(Cmb_Circo) + ","
      requete = requete + Iifch(Cmb_reins.Text) + ","
      requete = requete + Iifch(Txt_adr1) + ","
      requete = requete + Iifch(Txt_adr2) + ","
      requete = requete + Iifch(Txt_CP) + ","
      requete = requete + Iifch(Txt_Ville) + ","
      requete = requete + Iifch(Txt_Tel) + "," + "null" + "," + "null" + ","
      requete = requete + Iifch(Txt_regmat.Text) + ","
      If DTP_DateMariage <> "" Then
          requete = requete + Iifch(DTP_DateMariage) + ","
      Else
          requete = requete + "NULL" + ","
      End If
      requete = requete + Iifch(Txt_LieuMariage) + ")"
      Workspaces(0).Databases(0).Execute requete, dbFailOnError
      
      'On met à jour la combo contenant la liste des dossiers
      Call ListeDossiers(frm_DossierGestion.Dat_dossiers, frm_DossierGestion.Cmb_NumDossier)
      frm_DossierGestion.Cmb_NumDossier.ListIndex = PresentDansListBox(frm_DossierGestion.Cmb_NumDossier, Txt_num)

    Case "MODIFICATION"
      MousePointer = vbHourglass
      requete = "UPDATE DOSS"
      requete = requete + " SET SITF_CODE=" + CodeEntreCotes(Cmb_SituationFamiliale.Text)
      requete = requete + ",COMM_INSEE=" + CodeEntreCotes(Cmb_Commune.Text)
      requete = requete + ",CANT_NUM=" + CodeEntreCotes(Cmb_Canton.Text)
      requete = requete + ",CIRC_NUM=" + CodeEntreCotes(Cmb_Circo.Text)
      requete = requete + ",DOSS_REINS=" + Iifch(Cmb_reins.Text)
      requete = requete + ",DOSS_ADR1=" + Iifch(Txt_adr1)
      requete = requete + ",DOSS_ADR2=" + Iifch(Txt_adr2)
      requete = requete + ",DOSS_CP=" + Iifch(Txt_CP)
      requete = requete + ",DOSS_VILLE=" + Iifch(Txt_Ville)
      requete = requete + ",DOSS_NUMTEL=" + Iifch(Txt_Tel)
      requete = requete + ",SITU_REGMAT=" + Iifch(Txt_regmat.Text)
      If DTP_DateMariage <> "" Then
          requete = requete + ",SITU_DMARIA=" + Iifch(DTP_DateMariage)
      Else
          requete = requete + ",SITU_DMARIA=" + "NULL"
      End If
      requete = requete + ",SITU_LMARIA=" + Iifch(Txt_LieuMariage)
      requete = requete + " WHERE DOSS_NUM=" + Iifch(Txt_num)
      Workspaces(0).Databases(0).Execute requete, dbFailOnError
            
    Case "SUPPRESSION"
      If MsgBox("Cette action supprimera tous les enregistrements concernant" + Chr(10) _
          + "les données sur le dossier mais AUSSI : " + Chr(10) _
          + " - les placements éventuellement déjà effectués," + Chr(10) _
          + " - les données sur les individus du foyer demandeur," + Chr(10) _
          + " - les données sur les eventuelles demandes formulées", vbYesNo + vbCritical, "ATTENTION !") = vbYes Then
        MousePointer = vbHourglass
        
        sql = "select dema_num from dema where doss_num=" + Iifch(Txt_num.Text)
        Set RsDema = Db.OpenRecordset(sql)
        If Not RsDema.EOF Then
          RsDema.MoveFirst
          While Not RsDema.EOF
            NumDema = RsDema.Fields(0)
            sql = "select ende_num from coac where dema_num=" + Iifent(NumDema)
            Set RsCoac = Db.OpenRecordset(sql)
            sql = "select ende_num from fade where dema_num=" + Iifent(NumDema)
            Set RsFade = Db.OpenRecordset(sql)
            
            If Not RsCoac.EOF Then
              RsCoac.MoveFirst
              While Not RsCoac.EOF
                NumEnf = RsCoac.Fields(0)
                requete = "delete * from orso where ende_num=" + Iifent(NumEnf)
                Workspaces(0).Databases(0).Execute requete, dbFailOnError
                requete = "delete * from coac where ende_num=" + Iifent(NumEnf)
                Workspaces(0).Databases(0).Execute requete, dbFailOnError
                requete = "delete * from ende where ende_num=" + Iifent(NumEnf)
                Workspaces(0).Databases(0).Execute requete, dbFailOnError
                RsCoac.Requery
              Wend
            End If
            If Not RsFade.EOF Then
              RsFade.MoveFirst
              While Not RsFade.EOF
                NumEnf = RsFade.Fields(0)
                requete = "delete * from orso where ende_num=" + Iifent(NumEnf)
                Workspaces(0).Databases(0).Execute requete, dbFailOnError
                requete = "delete * from fade where ende_num=" + Iifent(NumEnf)
                Workspaces(0).Databases(0).Execute requete, dbFailOnError
                requete = "delete * from ende where ende_num=" + Iifent(NumEnf)
                Workspaces(0).Databases(0).Execute requete, dbFailOnError
                RsFade.Requery
              Wend
            End If
            
            requete = "DELETE * FROM SUPL WHERE PLAC_NUM IN ( "
            requete = requete + "SELECT PLAC_NUM FROM PLAC WHERE DEMA_NUM = " + Iifent(NumDema) + ")"
            Workspaces(0).Databases(0).Execute requete, dbFailOnError
            requete = "DELETE * FROM ENPL WHERE PLAC_NUM IN ( "
            requete = requete + "SELECT PLAC_NUM FROM PLAC WHERE DEMA_NUM = " + Iifent(NumDema) + ")"
            Workspaces(0).Databases(0).Execute requete, dbFailOnError
            requete = "DELETE * FROM PLAC WHERE DEMA_NUM = " + Iifent(NumDema)
            Workspaces(0).Databases(0).Execute requete, dbFailOnError
            requete = "DELETE * FROM RENO WHERE DEMA_NUM =" + Iifent(NumDema)
            Workspaces(0).Databases(0).Execute requete, dbFailOnError
            
            requete = "DELETE * FROM DEMA WHERE DEMA_NUM = " + Iifent(NumDema)
            Workspaces(0).Databases(0).Execute requete, dbFailOnError
            RsDema.Requery
          Wend
        End If
            
        requete = "DELETE * FROM INDI WHERE DOSS_NUM='" & Txt_num & "'"
        Workspaces(0).Databases(0).Execute requete, dbFailOnError
                        
        requete = "DELETE * FROM DOSS WHERE DOSS_NUM = " + Iifch(Txt_num)
        Workspaces(0).Databases(0).Execute requete, dbFailOnError

        Call ListeDossiers(frm_DossierGestion.Dat_dossiers, frm_DossierGestion.Cmb_NumDossier)
        frm_DossierGestion.Txt_adr1 = ""
        frm_DossierGestion.Txt_adr2 = ""
        frm_DossierGestion.Txt_Canton = ""
        frm_DossierGestion.Txt_Commune = ""
        frm_DossierGestion.Txt_Circonscription = ""
        frm_DossierGestion.Txt_CP = ""
        frm_DossierGestion.Txt_LieuMariage = ""
        frm_DossierGestion.Txt_RegimeMatrimonial = ""
        frm_DossierGestion.Txt_reinstruction = ""
        frm_DossierGestion.Txt_SituationFamiliale = ""
        frm_DossierGestion.Txt_Tel = ""
        frm_DossierGestion.Txt_Ville = ""
        frm_DossierGestion.DTP_DateMariage = "31/12/9999"
        frm_DossierGestion.Txt_Departement = "22"
        frm_DossierGestion.Txt_AnneeMoyenne = ""
        frm_DossierGestion.Txt_MoisMoyenne = ""
        
        frm_DossierGestion.Dat_demandes.RecordSource = "select * from dema where 1=2"
        frm_DossierGestion.Dat_demandes.Refresh
      End If
      
    End Select
    MousePointer = vbDefault
    Unload Me
    frm_DossierGestion.Dat_dossiers.Refresh
    Exit Sub

Err_cmd_Valider_Click:
Dim pbooAncienneGestionErreurs As Boolean
Dim pintTraitement As MessagesActions
'Affiche le code d'erreur approprié
    Select Case Err
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case 13:            pbooAncienneGestionErreurs = True: MsgBox "Valeur incompatible avec le type de colonne", vbOKOnly + vbExclamation
        Case 3022:          pbooAncienneGestionErreurs = True: MsgBox "Doublon" + ", création impossible", vbOKOnly + vbExclamation
        Case 3200:          pbooAncienneGestionErreurs = True: MsgBox "Reference" + ", suppression impossible", vbOKOnly + vbExclamation
        Case 3075:          pbooAncienneGestionErreurs = True: MsgBox "L'un des champs contient une cote, mise à jour impossible", vbOKOnly + vbExclamation
        Case 3315:          pbooAncienneGestionErreurs = True: MsgBox "Un champ obligatoire n'est pas renseigné, mise à jour impossible", vbOKOnly + vbExclamation
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
'   S'il s'agit de l'ancienne gestion d'erreurs
    If pbooAncienneGestionErreurs Then
        MousePointer = vbDefault
        Exit Sub
'   Sinon
    Else
        If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    End If

End Sub

Private Sub DTP_DateMariage_GotFocus()
    If (stb_DossiersFiche.Panels(4).Text = "CREATION") Or (stb_DossiersFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DossiersFiche.SimpleText = "Entrez la date de mariage du foyer demandeur (s'il y a lieu)"
    End If
End Sub

Private Sub Form_Load()

    QuitterCroix = True
    Txt_Departement = "22"
    Dat_ReqSituationFamiliale.DatabaseName = CheminBase
    Dat_Canton.DatabaseName = CheminBase
    Dat_Commune.DatabaseName = CheminBase
    Dat_Circonscriptions.DatabaseName = CheminBase
    
    Select Case frm_DossierGestion.stb_GestionDossier.Panels(4).Text
        Case "CREATION"
            stb_DossiersFiche.Panels(4) = "CREATION"
            ListeCantons Cmb_Canton
            ListeCirconscriptions Cmb_Circo
            ListeCommunes Cmb_Commune
            Cmb_Commune.ListIndex = 0
            ListeReinstruction Cmb_reins
            Cmb_reins.ListIndex = 0
            ListeSituationsFamiliales Cmb_SituationFamiliale
            Cmb_SituationFamiliale.ListIndex = 0
            DTP_DateMariage = Date
            DTP_DateMariage = ""
            
        Case "MODIFICATION"
            stb_DossiersFiche.Panels(4).Text = "MODIFICATION"
            ListeSituationsFamiliales Cmb_SituationFamiliale
            ListeReinstruction Cmb_reins
            ListeCantons Cmb_Canton
            ListeCommunes Cmb_Commune
            ListeCirconscriptions Cmb_Circo
            Cmb_Commune.ListIndex = 0

            
        Case "SUPPRESSION"
            stb_DossiersFiche.Panels(4).Text = "SUPPRESSION"
            ListeSituationsFamiliales Cmb_SituationFamiliale
            ListeReinstruction Cmb_reins
            ListeCantons Cmb_Canton
            ListeCommunes Cmb_Commune
            ListeCantons Cmb_Canton
            ListeCirconscriptions Cmb_Circo
           
    End Select
    
End Sub



Private Sub txt_adr1_gotfocus()
    If (stb_DossiersFiche.Panels(4).Text = "CREATION") Or (stb_DossiersFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DossiersFiche.SimpleText = "Entrez l'adresse du foyer demandeur"
    End If
End Sub

Private Sub txt_adr2_Gotfocus()
    If (stb_DossiersFiche.Panels(4).Text = "CREATION") Or (stb_DossiersFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DossiersFiche.SimpleText = "Entrez l'adresse du foyer demandeur"
    End If
End Sub



Private Sub txt_cp_Gotfocus()
    If (stb_DossiersFiche.Panels(4).Text = "CREATION") Or (stb_DossiersFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DossiersFiche.SimpleText = "Entrez le code postal du foyer demandeur"
    End If
End Sub

Private Sub Txt_LieuMariage_Gotfocus()
    If (stb_DossiersFiche.Panels(4).Text = "CREATION") Or (stb_DossiersFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DossiersFiche.SimpleText = "Entrez le lieu de mariage du foyer demandeur (s'il y a lieu)"
    End If
End Sub

Private Sub Txt_num_Gotfocus()
    If (stb_DossiersFiche.Panels(4).Text = "CREATION") Or (stb_DossiersFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DossiersFiche.SimpleText = "Entrez le numéro du dossier"
    End If
End Sub

Private Sub Txt_regmat_GotFocus()

  If (stb_DossiersFiche.Panels(4).Text = "CREATION") Or (stb_DossiersFiche.Panels(4).Text = "MODIFICATION") Then
    stb_DossiersFiche.SimpleText = "Entrez le régime matrimonial du foyer demandeur (s'il y a lieu)"
  End If
    
End Sub


Private Sub Txt_Tel_Gotfocus()
    If (stb_DossiersFiche.Panels(4).Text = "CREATION") Or (stb_DossiersFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DossiersFiche.SimpleText = "Entrez le numéro de téléphone du foyer demandeur"
    End If
End Sub

Private Sub txt_Ville_Gotfocus()
    If (stb_DossiersFiche.Panels(4).Text = "CREATION") Or (stb_DossiersFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DossiersFiche.SimpleText = "Entrez la ville du foyer demandeur"
    End If
End Sub

Public Sub ListeReinstruction(Boite As ComboBox)
'Permet de choisir un dossier en réinstruction ou non
'grâce à la Combo
        Boite.AddItem ("N")
        Boite.AddItem ("O")
End Sub

Public Sub ListeSituationsFamiliales(Boite As ComboBox)
'Place la liste des situations familiales dans la combo
    With Dat_ReqSituationFamiliale
        .RecordSource = "SELECT SITF_CODE, SITF_DESIG FROM SITF ORDER BY SITF_DESIG"
        .Refresh
        Boite.Clear
        While Not Dat_ReqSituationFamiliale.Recordset.EOF
            Boite.AddItem (.Recordset.Fields("SITF_DESIG") + " <" + CStr(.Recordset.Fields("SITF_CODE")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub


Public Sub ListeCantons(Boite As ComboBox)
'Place la liste des cantons dans la combo

    With Dat_Canton
        .RecordSource = "SELECT CANT_NOM, CANT_NUM FROM CANT ORDER BY CANT_NOM"
        .Refresh
        Boite.Clear
        While Not Dat_Canton.Recordset.EOF
            Boite.AddItem (.Recordset.Fields("CANT_NOM") + " <" + CStr(.Recordset.Fields("CANT_NUM")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub

Public Sub ListeCommunes(Boite As ComboBox)
'Place la liste des communes du département dans la combo
    With Dat_Commune
        .RecordSource = "SELECT COMM_NOM, COMM_INSEE FROM COMM ORDER BY COMM_NOM"
        .Refresh
        Boite.Clear
        While Not Dat_Commune.Recordset.EOF
            Boite.AddItem (.Recordset.Fields("COMM_NOM") + " <" + CStr(.Recordset.Fields("COMM_INSEE")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub



Public Sub ListeCirconscriptions(Boite As ComboBox)
'Place la liste des circonscriptions dans la combo
    With Dat_Circonscriptions
        .RecordSource = "SELECT CIRC_NOM, CIRC_NUM FROM CIRC ORDER BY CIRC_NOM"
        .Refresh
        Boite.Clear
        While Not Dat_Circonscriptions.Recordset.EOF
            Boite.AddItem (.Recordset.Fields("CIRC_NOM") + " <" + CStr(.Recordset.Fields("CIRC_NUM")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub



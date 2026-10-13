VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_TS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intervenants"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_TS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Dat_TS 
      Caption         =   "Dat_TS"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   9255
      Begin VB.ComboBox Cmb_Actif 
         Height          =   315
         ItemData        =   "frm_TS.frx":0442
         Left            =   6480
         List            =   "frm_TS.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Cmd_Actif 
         Caption         =   "En activité"
         Height          =   255
         Left            =   4920
         TabIndex        =   21
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox Cmb_Perm 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1200
         Width           =   3015
      End
      Begin VB.CommandButton cmd_ts_Perm 
         Caption         =   "&Permanence"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox Cmb_Qual 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox Cmb_Inst 
         Height          =   315
         ItemData        =   "frm_TS.frx":0446
         Left            =   6480
         List            =   "frm_TS.frx":0448
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cmd_ts_qua 
         Caption         =   "&Qualité"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmd_ts_cinst 
         Caption         =   "Code &Institution"
         Height          =   255
         Left            =   4920
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmd_ts_code 
         Caption         =   "&N° TS"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_ts_nom 
         Caption         =   "N&om"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txt_ts_code 
         Height          =   285
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   12
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txt_ts_nom 
         Height          =   285
         Left            =   6480
         MaxLength       =   30
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton Cmd_Annuler 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   7320
         Picture         =   "frm_TS.frx":044A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Annuler"
         Top             =   150
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "frm_TS.frx":054C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modifier les données de l'intervenant social sélectionné"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "frm_TS.frx":064E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Créer un nouvel intervenant social"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "frm_TS.frx":0750
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprimer l'intervenant social sélectionné"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Excel 
         Height          =   375
         Left            =   4560
         Picture         =   "frm_TS.frx":0852
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Export EXCEL"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8760
         Picture         =   "frm_TS.frx":09DC
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         Height          =   375
         Left            =   8040
         Picture         =   "frm_TS.frx":0ADE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Valider"
         Top             =   150
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmd_Actualiser 
         Default         =   -1  'True
         Height          =   375
         Left            =   6600
         Picture         =   "frm_TS.frx":0BE0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Rechercher d'après les critères"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_TS 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   17
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
            TextSave        =   "08/07/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "14:18"
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
   Begin MSDBGrid.DBGrid grd_TS 
      Bindings        =   "frm_TS.frx":0CE2
      Height          =   4095
      Left            =   1200
      OleObjectBlob   =   "frm_TS.frx":0CF7
      TabIndex        =   11
      Top             =   2400
      Width           =   7215
   End
End
Attribute VB_Name = "frm_TS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Validation(pvarIntervenantNuméro As Variant, pstrIntervenantNom As String)

Public Sub ListeActif(Boite As ComboBox)

  Boite.AddItem ("<Tous>")
  Boite.AddItem ("Oui")
  Boite.AddItem ("Non")

End Sub


Private Sub Cmd_Actif_Click()
  Cmb_Actif.ListIndex = 0
  Cmb_Actif.SetFocus
End Sub

Private Sub Cmd_Actualiser_Click()
Const pkstrNomProcédure As String = "Cmd_Actualiser_Click"
On Error GoTo Err_Cmd_Actualiser_Click

  Dim sel_actif, sel_code, sel_inst As String
  Dim sel_nom, sel_telprof, sel_qual, sel_perm As String
  Dim instSEARCH, qualSEARCH, permSEARCH As String
    
    Me.MousePointer = vbHourglass
    
    If Cmb_Actif.Text = "Oui" Then
        sel_actif = " AND TRSO_ACTIF=1"
    Else
      If Cmb_Actif.Text = "Non" Then
        sel_actif = " AND TRSO_ACTIF=0"
      End If
    End If
    
    If Cmb_Inst.ListIndex <> 0 Then
        instSEARCH = CodeEntreCotes(Cmb_Inst.Text)
    Else
        instSEARCH = ""
    End If
    If Cmb_Qual.ListIndex <> 0 Then
        qualSEARCH = CodeEntreCotes(Cmb_Qual.Text)
    Else
        qualSEARCH = ""
    End If
    If Cmb_Perm.ListIndex <> 0 Then
        permSEARCH = CodeEntreCotes(Cmb_Perm.Text)
    Else
        permSEARCH = ""
    End If
        
    If txt_ts_code <> "" Then sel_code = " AND TRSO_NUM like """ & txt_ts_code & "*"""
    If txt_ts_nom <> "" Then sel_nom = " AND TRSO_NOM like """ & txt_ts_nom & "*"""
    If instSEARCH <> "" Then sel_inst = " AND INST_CODE like """ & instSEARCH & "*"""
    If qualSEARCH <> "" Then sel_qual = " AND QUAL_CODE like """ & qualSEARCH & "*"""
    If permSEARCH <> "" Then sel_perm = " AND PERM_CODE like """ & permSEARCH & "*"""
    Dat_TS.RecordSource = "SELECT * FROM TRSO WHERE TRUE" + _
        sel_actif + sel_code + sel_nom + sel_inst + sel_telprof + sel_qual + sel_perm
        
    Dat_TS.Refresh
    Me.MousePointer = vbDefault
    
    Exit Sub
    
Err_Cmd_Actualiser_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        

End Sub

Private Sub cmd_Annuler_Click()
Const pkstrNomProcédure As String = "cmd_Annuler_Click"
On Error GoTo Err_cmd_Annuler_Click
        
'   Publier la suppression
    RaiseEvent Validation(Null, "")
    
    Unload Me

    Exit Sub
    
Err_cmd_Annuler_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
End Sub

Private Sub cmd_Creer_Click()

    stb_TS.Panels(4).Text = "CREATION"
    Load frm_TSFiche
    frm_TSFiche.Caption = "Création d'un nouvel intervenant social"
    frm_TSFiche.Txt_code = ""
    frm_TSFiche.Txt_Nom = ""
    frm_TSFiche.txt_telprof = ""
    
    With frm_TSFiche
        ListeInstitutions .Cmb_Inst
        .Cmb_Inst.RemoveItem (0)
        .Cmb_Inst.ListIndex = 0
    End With
    With frm_TSFiche
        ListeQualites .Cmb_Qual
        .Cmb_Qual.RemoveItem (0)
        .Cmb_Qual.ListIndex = 0
    End With
    With frm_TSFiche
        ListePermanences .Cmb_Perm
        .Cmb_Perm.RemoveItem (0)
        .Cmb_Perm.ListIndex = 0
    End With
    With frm_TSFiche
        ListeTitres .Cmb_Titr
        .Cmb_Titr.ListIndex = 0
    End With
    stb_TS.Panels(3).Text = ""
    frm_TSFiche.Show vbModal
  
    stb_TS.Panels(4).Text = ""
    
End Sub

Private Sub cmd_Excel_Click()

  ExportExcel (Dat_TS.RecordSource)

End Sub

Private Sub cmd_Modifier_Click()
    If Not Dat_TS.Recordset.EOF Then
        stb_TS.Panels(4).Text = "MODIFICATION"
        Load frm_TSFiche
        frm_TSFiche.Caption = "Modification d'un intervenant social"
        frm_TSFiche.Txt_code.Enabled = False
        frm_TSFiche.Txt_code = Dat_TS.Recordset("TRSO_NUM")
        
        If Dat_TS.Recordset("TRSO_TITR") <> "" Then
            frm_TSFiche.Cmb_Titr.AddItem (Dat_TS.Recordset("TRSO_TITR"))
        Else
            frm_TSFiche.Cmb_Titr.Text = ""
        End If
        If Dat_TS.Recordset("TRSO_NOM") <> "" Then
            frm_TSFiche.Txt_Nom = Dat_TS.Recordset("TRSO_NOM")
        Else
            frm_TSFiche.Txt_Nom = ""
        End If
        If Dat_TS.Recordset("TRSO_TELPROF") <> "" Then
            frm_TSFiche.txt_telprof = Dat_TS.Recordset("TRSO_TELPROF")
        Else
            frm_TSFiche.txt_telprof = ""
        End If
        If Dat_TS.Recordset("TRSO_ACTIF") = 1 Then
                frm_TSFiche.chk_actif.Value = 1
        Else
                frm_TSFiche.chk_actif.Value = 0
        End If
        With frm_TSFiche
            ListeQualites .Cmb_Qual
            .Cmb_Qual.RemoveItem (0)
            ListeInstitutions .Cmb_Inst
            .Cmb_Inst.RemoveItem (0)
            ListePermanences .Cmb_Perm
            .Cmb_Perm.RemoveItem (0)
            ListeTitres .Cmb_Titr
            .Cmb_Qual.ListIndex = PresentDansListBox(.Cmb_Qual, "<" + CStr(Dat_TS.Recordset.Fields("QUAL_CODE")) + ">")
            .Cmb_Inst.ListIndex = PresentDansListBox(.Cmb_Inst, "<" + CStr(Dat_TS.Recordset.Fields("INST_CODE")) + ">")
            .Cmb_Perm.ListIndex = PresentDansListBox(.Cmb_Perm, "<" + CStr(Dat_TS.Recordset.Fields("PERM_CODE")) + ">")
            .Cmb_Titr.ListIndex = PresentDansListBox(.Cmb_Titr, "")
        End With
        frm_TSFiche.Show vbModal
        stb_TS.Panels(4).Text = ""
    End If
End Sub

Private Sub cmd_Quitter_Click()
    Unload frm_TS
End Sub

Private Sub cmd_Supprimer_Click()
  If Not Dat_TS.Recordset.EOF Then
    stb_TS.Panels(4).Text = "SUPPRESSION"
    Load frm_TSFiche
    frm_TSFiche.Caption = "Suppression d'un intervenant social"
    frm_TSFiche.Fra_ts.Enabled = False
    
    frm_TSFiche.Txt_code.Enabled = False
    frm_TSFiche.Txt_code = Dat_TS.Recordset("TRSO_NUM")
    If Dat_TS.Recordset("TRSO_TITR") <> "" Then
        frm_TSFiche.Cmb_Titr.AddItem (Dat_TS.Recordset("TRSO_TITR"))
    Else
        frm_TSFiche.Cmb_Titr.Text = ""
    End If
    If Dat_TS.Recordset("TRSO_NOM") <> "" Then
        frm_TSFiche.Txt_Nom = Dat_TS.Recordset("TRSO_NOM")
    Else
        frm_TSFiche.Txt_Nom = ""
    End If
    If Dat_TS.Recordset("TRSO_TELPROF") <> "" Then
        frm_TSFiche.txt_telprof = Dat_TS.Recordset("TRSO_TELPROF")
    Else
        frm_TSFiche.txt_telprof = ""
    End If
    If Dat_TS.Recordset("TRSO_ACTIF") = 1 Then
        frm_TSFiche.chk_actif.Value = 1
    Else
        frm_TSFiche.chk_actif.Value = 0
    End If
    With frm_TSFiche
        ListeQualites .Cmb_Qual
        ListeInstitutions .Cmb_Inst
        ListePermanences .Cmb_Perm
        ListeTitres .Cmb_Titr
        .Cmb_Qual.ListIndex = PresentDansListBox(.Cmb_Qual, "<" + CStr(Dat_TS.Recordset.Fields("QUAL_CODE")) + ">")
        .Cmb_Inst.ListIndex = PresentDansListBox(.Cmb_Inst, "<" + CStr(Dat_TS.Recordset.Fields("INST_CODE")) + ">")
        .Cmb_Perm.ListIndex = PresentDansListBox(.Cmb_Perm, "<" + CStr(Dat_TS.Recordset.Fields("PERM_CODE")) + ">")
        .Cmb_Titr.ListIndex = PresentDansListBox(.Cmb_Titr, "")
    End With
    frm_TSFiche.Show vbModal
    stb_TS.Panels(4).Text = ""
  End If
End Sub

Private Sub cmd_ts_cinst_Click()
  Cmb_Inst.ListIndex = 0
  Cmb_Inst.SetFocus
End Sub

Private Sub cmd_ts_code_Click()
    txt_ts_code = ""
    txt_ts_code.SetFocus
End Sub

Private Sub cmd_ts_nom_Click()
    txt_ts_nom = ""
    txt_ts_nom.SetFocus
End Sub

Private Sub cmd_ts_Perm_Click()
    Cmb_Perm.ListIndex = 0
    Cmb_Perm.SetFocus
End Sub

Private Sub cmd_ts_qua_Click()
    Cmb_Qual.ListIndex = 0
    Cmb_Qual.SetFocus
End Sub

Private Sub cmd_Valider_Click()
Const pkstrNomProcédure As String = "cmd_Valider_Click"
On Error GoTo Err_cmd_Valider_Click

'   Publier la sélection
    RaiseEvent Validation(Dat_TS.Recordset("trso_num"), Dat_TS.Recordset("trso_nom"))
    
    Unload Me

    Exit Sub
    
Err_cmd_Valider_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
  
End Sub

Private Sub grd_TS_HeadClick(ByVal ColIndex As Integer)
  
  Dim colonne As String
  Dim classem As String
  
  colonne = grd_TS.Columns(ColIndex).DataField
  Call Tri(Dat_TS, grd_TS, colonne)
      
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_TS.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_TS.Panels(4).Text = "Classement " & classem
    
End Sub

Public Sub ListeInstitutions(Boite As ComboBox)
'Place la liste des institutions dans la combo
    With frm_TSFiche.Dat_ReqInst
        .RecordSource = "SELECT INST_CODE, INST_LIB FROM INST ORDER BY INST_LIB"
        .Refresh
        Boite.Clear
        Boite.AddItem ("<Tous>")
        While Not .Recordset.EOF
            Boite.AddItem (.Recordset.Fields("INST_LIB") + " <" + CStr(.Recordset.Fields("INST_CODE")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub

Public Sub ListeQualites(Boite As ComboBox)
'Place la liste des qualités dans la combo
    With frm_TSFiche.Dat_ReqQual
        .RecordSource = "SELECT QUAL_CODE, QUAL_LIB FROM QUAL ORDER BY QUAL_LIB"
        .Refresh
        Boite.Clear
        Boite.AddItem ("<Tous>")
        While Not .Recordset.EOF
            Boite.AddItem (.Recordset.Fields("QUAL_LIB") + " <" + CStr(.Recordset.Fields("QUAL_CODE")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub



Public Sub ListePermanences(Boite As ComboBox)
'Place la liste des permanences dans la combo
    With frm_TSFiche.Dat_ReqPerm
        .RecordSource = "SELECT PERM_CODE, PERM_LIEU, PERM_VILLE FROM PERM ORDER BY PERM_LIEU,PERM_VILLE"
        .Refresh
        Boite.Clear
        Boite.AddItem ("<Tous>")
        While Not .Recordset.EOF
            If .Recordset.Fields("PERM_LIEU") <> "Null" Then
                If .Recordset.Fields("PERM_VILLE") <> "Null" Then
                    Boite.AddItem (.Recordset.Fields("PERM_LIEU") + " " + .Recordset.Fields("PERM_VILLE") + " <" + CStr(.Recordset.Fields("PERM_CODE")) + ">")
                Else
                    Boite.AddItem (.Recordset.Fields("PERM_LIEU") + " <" + CStr(.Recordset.Fields("PERM_CODE")) + ">")
                End If
            Else
                If .Recordset.Fields("PERM_VILLE") <> "Null" Then
                    Boite.AddItem (.Recordset.Fields("PERM_LIEU") + " <" + CStr(.Recordset.Fields("PERM_CODE")) + ">")
                Else
                    Boite.AddItem ("<" + CStr(.Recordset("PERM_CODE")) + ">")
                End If
            End If
            .Recordset.MoveNext
        Wend
    End With
End Sub

Public Sub ListeTitres(Boite As ComboBox)
'Place la liste des titres dans la combo
        Boite.AddItem ("Mr")
        Boite.AddItem ("Mme")
        Boite.AddItem ("Mlle")
        Boite.AddItem ("Dr")
End Sub


'Initialiser la fenêtre
'   pbooGérerIntervenants:  booléen précisant s'il s'agit de gérer les intervenants [facultatif]
'   pbooAfficherAnnuler:    booléen précisant si le bouton Annuler doit apparaître  [facultatif]
'   pbooAfficherValider:    booléen précisant si le bouton Valider doit apparaître  [facultatif]

Public Sub Initialiser(Optional pbooGérerIntervenants As Boolean = True, Optional pbooAfficherAnnuler As Boolean = False, Optional pbooAfficherValider As Boolean = False)
Const pkstrNomProcédure As String = "Initialiser"
On Error GoTo Err_Initialiser
    
  With Dat_TS
    .DatabaseName = CheminBase
    .RecordSource = "select * from trso"
    .Refresh
  End With
  
  Cmd_Actualiser_Click
  ListeInstitutions Cmb_Inst
  Cmb_Inst.ListIndex = 0
  ListeQualites Cmb_Qual
  Cmb_Qual.ListIndex = 0
  ListePermanences Cmb_Perm
  Cmb_Perm.ListIndex = 0
  ListeActif Cmb_Actif
  Cmb_Actif.ListIndex = 0
  
  frm_TSFiche.QuitterCroix = False
  Unload frm_TSFiche
  
    cmd_Creer.Visible = pbooGérerIntervenants
    cmd_Modifier.Visible = pbooGérerIntervenants
    cmd_Supprimer.Visible = pbooGérerIntervenants
    cmd_Excel.Visible = pbooGérerIntervenants
    cmd_Quitter.Visible = pbooGérerIntervenants
        
    cmd_Annuler.Visible = pbooAfficherAnnuler
    cmd_Valider.Visible = pbooAfficherValider
        
    Exit Sub
    
Err_Initialiser:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SourcesDonnées_fiche 
   Caption         =   "Source de données fiche"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   ControlBox      =   0   'False
   Icon            =   "SourcesDonnées_Fiche.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8820
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Dat_wwhe 
      Caption         =   "Dat_wwhe"
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
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Cmd_Conditions 
      Caption         =   "Conditions de sélection"
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Data Dat_wsel 
      Caption         =   "Dat_wsel"
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
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox Cmb_Vue 
      Height          =   315
      ItemData        =   "SourcesDonnées_Fiche.frx":0442
      Left            =   3000
      List            =   "SourcesDonnées_Fiche.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Liste des vues existantes"
      Top             =   2280
      Width           =   2775
   End
   Begin VB.ComboBox Cmb_Bloque 
      Height          =   315
      ItemData        =   "SourcesDonnées_Fiche.frx":0446
      Left            =   3000
      List            =   "SourcesDonnées_Fiche.frx":0450
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "La source de données est-elle bloquée ?"
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Cmd_ordre 
      Caption         =   "Choix de l'ordre"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton Cmd_colonnes 
      Caption         =   "Choix des colonnes"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Data Dat_id 
      Caption         =   "Dat_id"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Dat_vue 
      Caption         =   "Dat_vue"
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
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Txt_nom 
      Height          =   285
      Left            =   3000
      MaxLength       =   30
      TabIndex        =   2
      ToolTipText     =   "Nom de la source de données"
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Txt_id 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      ToolTipText     =   "Identifiant de la source de données"
      Top             =   1320
      Width           =   615
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   8535
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   7800
         Picture         =   "SourcesDonnées_Fiche.frx":045E
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Annuler 
         Height          =   375
         Left            =   6360
         Picture         =   "SourcesDonnées_Fiche.frx":0560
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   7080
         Picture         =   "SourcesDonnées_Fiche.frx":0662
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_FicheSource 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5295
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12435
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Fixe"
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Vue"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Nom"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Identifiant"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1320
      Width           =   735
   End
End
Attribute VB_Name = "SourcesDonnées_fiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents RsColonne As ADODB.Recordset
Attribute RsColonne.VB_VarHelpID = -1
Public WithEvents RsSelect As ADODB.Recordset
Attribute RsSelect.VB_VarHelpID = -1

Public ChoixFait As Boolean
Public SelectFait As Boolean


Sub InitialisationRs(IdentVue As Integer)
Const pkstrNomProcédure As String = "InitialisationRs"
On Error GoTo Err_InitialisationRs

  Dim curseur As Integer
  
  Set RsColonne = New ADODB.Recordset
    
  With RsColonne
    .Fields.Append "IdentSource", adInteger
    .Fields.Append "NumOrdre", adInteger
    .Fields.Append "NomColonne", adVarChar, daoWCOL_NOM
    .Fields.Append "Aliascolonne", adVarChar, 30, adFldMayBeNull
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
  End With
  
  For curseur = 0 To Db.QueryDefs(IdentVue).Fields.Count - 1
    
    RsColonne.AddNew
    RsColonne!identsource = Txt_id
    RsColonne!NumOrdre = curseur + 1
    RsColonne!nomcolonne = Db.QueryDefs(IdentVue).Fields(curseur).Name
    'RsColonne!Aliascolonne = ""
    RsColonne.Update
  
  Next curseur
  
  Set Colonnes_Choix.Grd_Colonne.DataSource = RsColonne

    Exit Sub
    
Err_InitialisationRs:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub




Sub InitialisationSelect()
Const pkstrNomProcédure As String = "InitialisationSelect"
On Error GoTo Err_InitialisationSelect

  Dim curseur As Integer
  
  Set RsSelect = New ADODB.Recordset
    
  With RsSelect
    .Fields.Append "NomColonne", adVarChar, daoWCOL_NOM   'cette colonne sera invisible dans le datagrid5
    .Fields.Append "Alias", adVarChar, 30
    .Fields.Append "Condition", adVarChar, 30, adFldMayBeNull
    .Fields.Append "BorneInf", adVarChar, 30, adFldMayBeNull
    .Fields.Append "BorneSup", adVarChar, 30, adFldMayBeNull
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
'       Trier sur l'alias
'           Note: Le tri doit être réalisé après l'ouverture du Recordset, au risque de générer l'erreur -2147467259 ("Un état E_FAIL a été renvoyé par le fournisseur de données ou par un autre service.")
        .Sort = "Alias ASC"
  End With
  
  Dat_wsel.Recordset.MoveFirst
  While Not Dat_wsel.Recordset.EOF
    RsSelect.AddNew
    RsSelect!nomcolonne = Dat_wsel.Recordset("wcol_nom")
    RsSelect!Alias = Dat_wsel.Recordset("wsel_alias")
    RsSelect.Update
    Dat_wsel.Recordset.MoveNext
  Wend
  
  Set Colonnes_Select.Grd_select.DataSource = RsSelect

    Exit Sub
    
Err_InitialisationSelect:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


Sub ModificationRs()

Dim sql As String
Dim trouve As Boolean

  sql = "select * from wsel where wsou_id=" + Txt_id + " order by wsel_numordre"
  Dat_wsel.RecordSource = sql
  Dat_wsel.Refresh
  
  While Not Dat_wsel.Recordset.EOF
    RsColonne.MoveFirst
    trouve = False
    While Not RsColonne.EOF And Not trouve
      If Dat_wsel.Recordset("wsel_numordre") = RsColonne!NumOrdre Then
        trouve = True
      Else
        RsColonne.MoveNext
      End If
    Wend
    If trouve Then RsColonne!Aliascolonne = Dat_wsel.Recordset("wsel_alias")
    Dat_wsel.Recordset.MoveNext
  Wend
  
End Sub

Sub ModificationSelect()

Dim trouve As Boolean

  'vérifie s'il existe une base locale pour les conditions de sélection de cette source de données
  'DbAccessPath = "c:\Program Files\" + App.EXEName + "\" + Txt_id + ".mdb"
  'If Dir$(DbAccessPath, vbNormal) <> "" Then  'récupère les conditions de sélection
  '  Workspaces(0).OpenDatabase (DbAccessPath) 'situées dans la base locale
  '  With Dat_wwhe
  '    .DatabaseName = DbAccessPath
  '    .RecordSource = "select * from wwhe where wsou_id=" + CStr(Txt_id)
  '    .Refresh
  '    .Recordset.MoveFirst
  '  End With
  '  While Not Dat_wwhe.Recordset.EOF
  '    'met à jour le grid
  '    RsSelect.MoveFirst
  '    trouve = False
  '    While Not RsSelect.EOF And Not trouve
  '      If Dat_wwhe.Recordset("wcol_nom") = RsSelect(0) Then
  '        trouve = True
  '      Else
  '        RsSelect.MoveNext
  '      End If
   '   Wend
   '   If trouve Then
   ''     If Dat_wwhe.Recordset("wwhe_cond") <> "" Then RsSelect(2) = Dat_wwhe.Recordset("wwhe_cond")
   '     If Dat_wwhe.Recordset("wwhe_inf") <> "" Then RsSelect(3) = Dat_wwhe.Recordset("wwhe_inf")
   '     If Dat_wwhe.Recordset("wwhe_sup") <> "" Then RsSelect(4) = Dat_wwhe.Recordset("wwhe_sup")
   '   End If
   '   Dat_wwhe.Recordset.MoveNext
   ' Wend
  'Else  'récupére les conditions de sélection de la base centrale
    With Dat_wwhe
      .DatabaseName = CheminBase
      .RecordSource = "select * from wwhe where wsou_id=" + CStr(Txt_id)
      .Refresh
    End With
    While Not Dat_wwhe.Recordset.EOF
      'met à jour le grid
      RsSelect.MoveFirst
      trouve = False
      While Not RsSelect.EOF And Not trouve
        If Dat_wwhe.Recordset("wcol_nom") = RsSelect(0) Then
          trouve = True
        Else
          RsSelect.MoveNext
        End If
      Wend
      If trouve Then
        If Dat_wwhe.Recordset("wwhe_cond") <> "" Then RsSelect(2) = Dat_wwhe.Recordset("wwhe_cond")
        If Dat_wwhe.Recordset("wwhe_inf") <> "" Then RsSelect(3) = Dat_wwhe.Recordset("wwhe_inf")
        If Dat_wwhe.Recordset("wwhe_sup") <> "" Then RsSelect(4) = Dat_wwhe.Recordset("wwhe_sup")
      End If
      Dat_wwhe.Recordset.MoveNext
    Wend
  'End If

End Sub

Sub ParametreDefaut()

Dim trouve As Boolean

  With Dat_wwhe
    .DatabaseName = CheminBase
    .RecordSource = "select * from wwhe where wsou_id=" + CStr(Txt_id)
    .Refresh
  End With
  While Not Dat_wwhe.Recordset.EOF
    'met à jour le grid
    RsSelect.MoveFirst
    trouve = False
    While Not RsSelect.EOF And Not trouve
      If Dat_wwhe.Recordset("wcol_nom") = RsSelect(0) Then
        trouve = True
      Else
        RsSelect.MoveNext
      End If
    Wend
    If trouve Then
      If Dat_wwhe.Recordset("wwhe_cond") <> "" Then RsSelect(2) = Dat_wwhe.Recordset("wwhe_cond")
      If Dat_wwhe.Recordset("wwhe_inf") <> "" Then RsSelect(3) = Dat_wwhe.Recordset("wwhe_inf")
      If Dat_wwhe.Recordset("wwhe_sup") <> "" Then RsSelect(4) = Dat_wwhe.Recordset("wwhe_sup")
    End If
    Dat_wwhe.Recordset.MoveNext
  Wend

End Sub


Private Sub cmd_Annuler_Click()

Dim sql As String

  Select Case Caption
  
    Case "Création d'une source de données"
      sql = "DELETE * FROM WOBY WHERE WSOU_ID=" + Txt_id
      Workspaces(0).Databases(0).Execute sql, dbFailOnError
      sql = "DELETE * FROM WSEL WHERE WSOU_ID=" + Txt_id
      Workspaces(0).Databases(0).Execute sql, dbFailOnError
      sql = "DELETE * FROM WWHE WHERE WSOU_ID=" + Txt_id
      Workspaces(0).Databases(0).Execute sql, dbFailOnError
      sql = "DELETE * FROM WSOU WHERE WSOU_ID=" + Txt_id
      Workspaces(0).Databases(0).Execute sql, dbFailOnError
      
    Case "Modification d'une source de données"
      MsgBox "Les modifications apportées sur le choix et le tri des colonnes ne peuvent être annulées.", vbInformation, "Source de données"

  End Select
  
  Unload Me
End Sub

Private Sub Cmd_colonnes_Click()

Dim IdVue, j As Integer
Dim sql As String
  
  If Cmb_Vue.Text <> "" Then
    Load Colonnes_Choix
    IdVue = 0
    While Db.QueryDefs(IdVue).Name <> Cmb_Vue.Text
      IdVue = IdVue + 1
    Wend
  End If
  
  Select Case Caption
    Case "Création d'une source de données"
      InitialisationRs (IdVue)
      If ChoixFait Then
        ModificationRs
        Colonnes_Choix.stb_ChoixColonnes.Panels(4) = "MODIFICATION"
      Else
        Colonnes_Choix.stb_ChoixColonnes.Panels(4) = "CREATION"
      End If
      RsColonne.MoveFirst
    
    Case "Modification d'une source de données"
      Colonnes_Choix.stb_ChoixColonnes.Panels(4) = "MODIFICATION"
      InitialisationRs (IdVue)
      ModificationRs
      RsColonne.MoveFirst
      
  End Select
  
  Colonnes_Choix.Show vbModal
  
End Sub

Private Sub Cmd_Conditions_Click()

Dim sql As String
Dim vntTemp As Variant

  sql = "select * from wsel where wsou_id=" + CStr(Txt_id)
  With Dat_wsel
    .RecordSource = sql
    .Refresh
  End With
  If Dat_wsel.Recordset.EOF Then
    MsgBox "Vous ne pouvez pas effectuer de sélection." + Chr(13) + _
           "Cette source de données ne contient pas de colonnes !", vbExclamation, "ERREUR"
  Else
    Load Colonnes_Select
    Colonnes_Select.FeuilleAppelante = Me.Name
    Colonnes_Select.Vue = Cmb_Vue.Text
    Select Case Caption
      Case "Création d'une source de données"
        InitialisationSelect
        If SelectFait Then
          ModificationSelect
          Colonnes_Select.stb_ChoixColonnes.Panels(4).Text = "MODIFICATION"
        Else
          Colonnes_Select.stb_ChoixColonnes.Panels(4).Text = "CREATION"
        End If
        
      Case "Modification d'une source de données"
        Colonnes_Select.stb_ChoixColonnes.Panels(4).Text = "MODIFICATION"
        InitialisationSelect
        ModificationSelect
    
    End Select
    Colonnes_Select.Show vbModal
  End If
  
End Sub









Private Sub Cmd_ordre_Click()

Dim sql As String
Dim vntTemp As Variant

  sql = "select wcol_nom from wsel where wsou_id=" + CStr(Txt_id)
  With Dat_wsel
    .RecordSource = sql
    .Refresh
  End With
  If Dat_wsel.Recordset.EOF Then
    MsgBox "Vous ne pouvez pas effectuer de tri." + Chr(13) + _
           "Cette source de données ne contient pas de colonnes !", vbExclamation, "ERREUR"
  Else
    Load Colonnes_Tri
    'colonnes faisant déja partie du tri
    sql = "select woby.*,wsel.wsel_alias from woby,wsel where woby.wsou_id=" + CStr(Txt_id)
    sql = sql + " and woby.wsou_id=wsel.wsou_id and woby.wcol_nom=wsel.wcol_nom" + " and woby.wcol_nom in "
    sql = sql + "(select wcol_nom from woby where wsou_id=" + CStr(Txt_id) + ") order by woby_numordre"
    Colonnes_Tri.Dat_woby.RecordSource = sql
    Colonnes_Tri.Dat_woby.Refresh
    'remplissage de la liste des colonnes triées
    Do While Not Colonnes_Tri.Dat_woby.Recordset.EOF
      vntTemp = Colonnes_Tri.Dat_woby.Recordset("wsel_alias")
      If IsNull(vntTemp) Then vntTemp = ""
      Colonnes_Tri.Lst_ColonneTri.AddItem CStr(vntTemp)
      vntTemp = Colonnes_Tri.Dat_woby.Recordset("woby_sens")
      If IsNull(vntTemp) Then
        vntTemp = ""
      Else
        If vntTemp = "A" Then
          Colonnes_Tri.Lst_Sens.AddItem "Ascendant"
        Else
          Colonnes_Tri.Lst_Sens.AddItem "Descendant"
        End If
      End If
      Colonnes_Tri.Dat_woby.Recordset.MoveNext
    Loop
    
    'colonnes ne faisant pas partie du tri
    sql = "select * from wsel where wsou_id=" + CStr(Txt_id)
    sql = sql + " and wcol_nom not in (select wcol_nom from woby where wsou_id=" + CStr(Txt_id) + ")"
    Colonnes_Tri.Dat_wsel.RecordSource = sql
    Colonnes_Tri.Dat_wsel.Refresh
    'remplissage de la liste des colonnes "non triées"
    Do While Not Colonnes_Tri.Dat_wsel.Recordset.EOF
      'modifié le 18/05/2010 par Caroline
      vntTemp = Colonnes_Tri.Dat_wsel.Recordset("WSEL_ALIAS")
      'vntTemp = Colonnes_Tri.Dat_wsel.Recordset(3)
      If IsNull(vntTemp) Then vntTemp = ""
      Colonnes_Tri.Lst_colonnes.AddItem CStr(vntTemp)
      Colonnes_Tri.Dat_wsel.Recordset.MoveNext
    Loop
    
    Colonnes_Tri.Show vbModal
  End If

End Sub

Private Sub cmd_Quitter_Click()
  Unload Me
End Sub

Private Sub cmd_Valider_Click()

Dim sql As String
Dim bloque As Integer

On Error GoTo Rupture_Integrite

  Select Case Caption
    Case "Création d'une source de données"
      'on crée un nouvel enregistrement dans la table source de données
      If Cmb_Bloque.Text = "Oui" Then
        bloque = 1
      Else
        bloque = 0
      End If
      
      sql = "INSERT INTO WSOU VALUES("
      sql = sql + CStr(Txt_id) + ","
      sql = sql + Iifch(Txt_nom) + ","
      sql = sql + Iifch(Cmb_Vue.Text) + ","
      sql = sql + CStr(bloque) + ")"
      Workspaces(0).Databases(0).Execute sql, dbFailOnError
      Cmd_colonnes.Enabled = True
      Cmd_Conditions.Enabled = True
      Cmd_ordre.Enabled = True
    
    Case "Modification d'une source de données"
      sql = "UPDATE WSOU SET WSOU_NOM=" + Iifch(Txt_nom)
      If Cmb_Bloque.Text = "Oui" Then
        sql = sql + ",WSOU_BLOQUE=1"
      Else
        sql = sql + ",WSOU_BLOQUE=0"
      End If
      sql = sql + " WHERE WSOU_ID=" & Txt_id
      Workspaces(0).Databases(0).Execute sql, dbFailOnError
      
    Case "Suppression d'une source de données"
      sql = "DELETE * FROM WOBY WHERE WSOU_ID=" + Txt_id
      Workspaces(0).Databases(0).Execute sql, dbFailOnError
      sql = "DELETE * FROM WSEL WHERE WSOU_ID=" + Txt_id
      Workspaces(0).Databases(0).Execute sql, dbFailOnError
      sql = "DELETE * FROM WWHE WHERE WSOU_ID=" + Txt_id
      Workspaces(0).Databases(0).Execute sql, dbFailOnError
      sql = "DELETE * FROM WSOU WHERE WSOU_ID=" + Txt_id
      Workspaces(0).Databases(0).Execute sql, dbFailOnError

  End Select

  SourcesDonnées_liste.Dat_source.Refresh
  cmd_Annuler.Enabled = False
  cmd_Valider.Enabled = False
  'Unload Me
  Exit Sub
  
Rupture_Integrite:
'Affiche le code d'erreur approprié
  Select Case Err
    Case 13
      MsgBox "Valeur incompatible avec le type de colonne", vbOKOnly + vbExclamation
  
    Case 3022
      MsgBox "Doublon" + ", création impossible", vbOKOnly + vbExclamation
  
    Case 3200
      MsgBox "Reference" + ", suppression impossible", vbOKOnly + vbExclamation
  
    Case 3075
      MsgBox "L'un des champs contient une cote, mise à jour impossible", vbOKOnly + vbExclamation
  
    Case 3315
      MsgBox "Un champ obligatoire n'est pas renseigné, mise à jour impossible", vbOKOnly + vbExclamation
  
    Case Else
      MsgBox "Erreur" + CStr(Err) + " " + Err.Description + " survenue, mise à jour impossible", vbOKOnly + vbExclamation
  End Select
  Exit Sub
End Sub



Private Sub Form_Load()

Dim i As Integer

  ChoixFait = False   'permet de savoir lors d'une création si l'utilisateur est déjà
  SelectFait = False  'aller dans la fenetre Colonnes_choix pour choisir ses champs idem pour Colonnes_select

  Dat_wwhe.DatabaseName = CheminBase
  Dat_id.DatabaseName = CheminBase
  Dat_wsel.DatabaseName = CheminBase
  
  For i = 0 To Db.QueryDefs.Count - 1
    If Left(Db.QueryDefs(i).Name, 6) = "WMAIL_" Then
      Cmb_Vue.AddItem (Db.QueryDefs(i).Name)
    End If
  Next
  
End Sub



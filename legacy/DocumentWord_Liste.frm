VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DocumentWord_liste 
   Caption         =   "Liste des documents Word"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "DocumentWord_Liste.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   9255
      Begin VB.TextBox Txt_DocNom 
         Height          =   285
         Left            =   6480
         MaxLength       =   50
         TabIndex        =   16
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Cmd_DocNom 
         Caption         =   "Document Nom"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Txt_DocIdent 
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Cmd_DocIdent 
         Caption         =   "Document Ident"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Txt_Source 
         Height          =   285
         Left            =   6480
         TabIndex        =   12
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Cmd_Source 
         Caption         =   "Source Nom"
         Height          =   255
         Left            =   4920
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Txt_Famille 
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Cmd_Famille 
         Caption         =   "Famille Nom"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Data Dat_wdoc 
      Caption         =   "Dat_wdoc"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "DocumentWord_Liste.frx":0442
      Height          =   3495
      Left            =   120
      OleObjectBlob   =   "DocumentWord_Liste.frx":0459
      TabIndex        =   6
      Top             =   1920
      Width           =   9255
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton Cmd_Actualiser 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   7920
         Picture         =   "DocumentWord_Liste.frx":1354
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Actualiser la liste des documents d'après les critères de recherche"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton Cmd_Mailings 
         Height          =   375
         Left            =   3120
         Picture         =   "DocumentWord_Liste.frx":1456
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Préciser les conditions de sélection puis fusionner"
         Top             =   150
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Data Dat_interim 
         Caption         =   "Dat_interim"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   5280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   720
         Picture         =   "DocumentWord_Liste.frx":1760
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Modifier le document Word sélectionné"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Creer 
         Height          =   375
         Left            =   120
         Picture         =   "DocumentWord_Liste.frx":1862
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Créer un document Word"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Supprimer 
         Height          =   375
         Left            =   1320
         Picture         =   "DocumentWord_Liste.frx":1964
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Supprimer le document Word sélectionné"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   8520
         Picture         =   "DocumentWord_Liste.frx":1A66
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fermer la fenêtre"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_ListeDoc 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   5610
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
            TextSave        =   "08/08/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15:38"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9128
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "DocumentWord_liste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents RsSelect As ADODB.Recordset
Attribute RsSelect.VB_VarHelpID = -1

Sub CreationBaseAccess()
'Création de la base de données résultat

  If Dir$(DbAccessResPath, vbNormal) <> "" Then
    Kill DbAccessResPath
  End If
  
  Set DbAccessRes = Workspaces(0).CreateDatabase(DbAccessResPath, dbLangGeneral, dbVersion30)
  
End Sub

Sub CreationFichierData()

Dim Fichier As String
Dim Enr As String

Dim j As Integer
Dim requete As String
Dim id As Index
Dim i As Integer
Dim sql, NomCol As String
Dim Alias As String
Dim Rs As Recordset
  
  'création d'un fichier TXT contenant les champs de fusion
  Fichier = FichierDonnées
  If Dir$(Fichier, vbNormal) <> "" Then
    Kill Fichier
  End If
  Open Fichier For Output As #1
  
  'La 1ère ligne contient l'alias du nom des champs
  Enr = ""
  
  'recherche l'alias de la colonne
  sql = "select * from wsel where wsou_id=" + CStr(Dat_wdoc.Recordset("wsou_id")) + " order by wcol_nom"
  Set Rs = Db.OpenRecordset(sql)
  Rs.MoveFirst
  For i = 0 To Dat_Interim.Recordset.Fields.Count - 1
    Alias = Rs("wsel_alias")
    Enr = Enr + Alias + Chr(9)
    Rs.MoveNext
  Next i
  Rs.Close
  'insère l'alias des champs dans le fichier
  Enr = Left(Enr, Len(Enr) - 1)
  Print #1, Enr
  
  'crée la requête permettant d'obtenir les enregistrements de la vue concernée
  sql = "select * from wsel where wsou_id=" + CStr(Dat_wdoc.Recordset("wsou_id")) + " order by wcol_nom"
  Set Rs = Db.OpenRecordset(sql)
  requete = "select "
  Rs.MoveFirst
  While Not Rs.EOF
    requete = requete + Rs("wcol_nom") + ","
    Rs.MoveNext
  Wend
  Rs.Close
  requete = Left(requete, Len(requete) - 1)
  requete = requete + " from " + Dat_wdoc.Recordset("wsou_vue")
  
  'crée et insère 1 enregistrement dans le fichier de données
  Set Rs = Db.OpenRecordset(requete)
  Enr = ""
  For j = 0 To Rs.Fields.Count - 1
    If IsNull(Rs.Fields(j)) Then
      Enr = Enr + "" + Chr(9)
    Else
      Enr = Enr + CStr(Rs.Fields(j)) + Chr(9)
    End If
  Next j
  Rs.Close
  Enr = Left(Enr, Len(Enr) - 1)
  Print #1, Enr
  
  Close #1
  
End Sub


Sub CreationFichierDataFusion()

Dim Fichier As String
Dim Enr As String

Dim j As Integer
Dim requete As String
Dim id As Index
Dim i As Integer
Dim sql, NomCol As String
Dim Alias As String
Dim Rs As Recordset
  
  'création d'un fichier TXT contenant les champs de fusion
  Fichier = FichierDonnées
  If Dir$(Fichier, vbNormal) <> "" Then
    Kill Fichier
  End If
  Open Fichier For Output As #1
  
  'La 1ère ligne contient l'alias du nom des champs
  Enr = ""
  
  'recherche l'alias de la colonne
  sql = "select * from wsel where wsou_id=" + CStr(Dat_wdoc.Recordset("wsou_id")) + " order by wcol_nom"
  Set Rs = Db.OpenRecordset(sql)
  Rs.MoveFirst
  For i = 0 To Dat_Interim.Recordset.Fields.Count - 1
    Alias = Rs("wsel_alias")
    Enr = Enr + Alias + Chr(9)
    Rs.MoveNext
  Next i
  Rs.Close
  'insère l'alias des champs dans le fichier
  Enr = Left(Enr, Len(Enr) - 1)
  Print #1, Enr
  
  'crée et insère les enregistrements dans le fichier de données
  Dat_Interim.Recordset.MoveFirst
  While Not Dat_Interim.Recordset.EOF
    Enr = ""
    For j = 0 To Dat_Interim.Recordset.Fields.Count - 1
      If IsNull(Dat_Interim.Recordset.Fields(j)) Then
        Enr = Enr + "" + Chr(9)
      Else
        Enr = Enr + CStr(Dat_Interim.Recordset.Fields(j)) + Chr(9)
      End If
    Next j
    Enr = Left(Enr, Len(Enr) - 1)
    Print #1, Enr
    Dat_Interim.Recordset.MoveNext
  Wend
  
  Close #1

End Sub


Sub CreationTableWRES()
'création de la table dans la nouvelle base

Dim td As TableDef
Dim fd As Field
Dim id As Index
Dim i As Integer
Dim sql, NomCol As String
Dim Alias As String
Dim RsAlias As Recordset
  
  'création de la table
  Set td = DbAccessRes.CreateTableDef("WRES")
  'création des champs de la table
  'recherche l'alias de la colonne
  sql = "select * from wsel where wsou_id=" + CStr(Dat_wdoc.Recordset("wsou_id")) + " order by wcol_nom"
  Set RsAlias = Db.OpenRecordset(sql)
  RsAlias.MoveFirst
  For i = 0 To Dat_Interim.Recordset.Fields.Count - 1
    Alias = RsAlias("wsel_alias")
    Set fd = td.CreateField(Alias, Dat_Interim.Recordset.Fields(i).Type, Dat_Interim.Recordset.Fields(i).Size)
    'ajout du champ à la collection des champs de la table
    td.Fields.Append fd
    RsAlias.MoveNext
  Next i
  RsAlias.Close
  'ajout de la table à la collection des tables
  DbAccessRes.TableDefs.Append td
  
End Sub


Sub CreerRequeteDoc()

Dim sql, requete As String

  sql = "select * from wsel where wsou_id=" + CStr(Dat_wdoc.Recordset("wsou_id")) + " order by wsel_numordre"
  With Dat_Interim
    .DatabaseName = CheminBase
    .RecordSource = sql
    .Refresh
    .Recordset.MoveFirst
  End With
  
  requete = "select "
 
  'colonnes
  While Not Dat_Interim.Recordset.EOF
    requete = requete + Dat_Interim.Recordset("wcol_nom") + ","
    Dat_Interim.Recordset.MoveNext
  Wend
  'enlève la dernière virgule
  requete = Left(requete, Len(requete) - 1)
  requete = requete + " from " + Dat_wdoc.Recordset("wsou_vue")
  
  Dat_Interim.RecordSource = requete
  Dat_Interim.Refresh

End Sub

Sub DefautRsSelect()

Dim trouve As Boolean

  Colonnes_Select.stb_ChoixColonnes.Panels(3).Text = "Base centrale"
  With Dat_Interim
    .DatabaseName = CheminBase
    .RecordSource = "select * from wwhe where wsou_id=" + CStr(Dat_wdoc.Recordset("wsou_id"))
    .Refresh
  End With
  While Not Dat_Interim.Recordset.EOF
    'met à jour le grid
    RsSelect.MoveFirst
    trouve = False
    While Not RsSelect.EOF And Not trouve
      If Dat_Interim.Recordset("wcol_nom") = RsSelect(0) Then
        trouve = True
      Else
        RsSelect.MoveNext
      End If
    Wend
    If trouve Then
      If Dat_Interim.Recordset("wwhe_cond") <> "" Then RsSelect(2) = Dat_Interim.Recordset("wwhe_cond")
      If Dat_Interim.Recordset("wwhe_inf") <> "" Then RsSelect(3) = Dat_Interim.Recordset("wwhe_inf")
      If Dat_Interim.Recordset("wwhe_sup") <> "" Then RsSelect(4) = Dat_Interim.Recordset("wwhe_sup")
    End If
    Dat_Interim.Recordset.MoveNext
  Wend

End Sub

Sub FusionDoc()

Dim myword As Object
Dim NomFichier As String
Dim CheminFichier
Dim fd As Field
Dim IdVue As Integer
Dim trouve As Boolean

  'recherche l'identifiant de la vue utilisée par la source de données
  trouve = False
  IdVue = 0
  While Not trouve And IdVue <= Db.QueryDefs.Count - 1
    If Db.QueryDefs(IdVue).Name = Dat_wdoc.Recordset("wsou_vue") Then
      trouve = True
    Else
      IdVue = IdVue + 1
    End If
  Wend
  If Not trouve Then
    MsgBox "La vue, sur laquelle s'appuie la source de données, n'existe pas.", vbExclamation, "ERREUR"
  Else
    CreerRequete (IdVue)
    If Dat_Interim.Recordset.EOF Then
      MsgBox "Aucun dossier ne répond aux conditions sélectionnées.", vbExclamation
    Else
      CreationFichierDataFusion
      
      NomFichier = Dat_wdoc.Recordset("wdoc_nom")

      If Right(CheminWord, 1) = "\" Then
        CheminFichier = CheminWord + NomFichier + ".doc"
      Else
        CheminFichier = CheminWord + "\" + NomFichier + ".doc"
      End If
      
      Set myword = GetObject(CheminFichier)
      myword.Parent.WindowState = wdWindowStateMaximize
      myword.Parent.DOCUMENTS(1).Activate
      myword.Parent.DOCUMENTS(1).Windows.Add
      With myword.Parent.ActiveDocument.MailMerge
        .MainDocumentType = wdFormLetters
        .OpenDataSource Name:=FichierDonnées, _
                        ReadOnly:=True, _
                        linktosource:=True
        .Destination = wdSendToNewDocument 'envoie le résultat du publipostage dans un nouveau document
        .Execute
      End With
      myword.Parent.Visible = True
      Set myword = Nothing
    End If
  End If
  
End Sub

Sub InitRsSelect()
Const pkstrNomProcédure As String = "InitRsSelect"
On Error GoTo Err_InitRsSelect

Dim curseur As Integer

  Set RsSelect = New ADODB.Recordset
  With RsSelect
    .Fields.Append "NomColonne", adVarChar, daoWCOL_NOM
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

  Dat_Interim.Recordset.MoveFirst
  While Not Dat_Interim.Recordset.EOF
    RsSelect.AddNew
    RsSelect!nomcolonne = Dat_Interim.Recordset("wcol_nom")
    RsSelect!Alias = Dat_Interim.Recordset("wsel_alias")
    RsSelect.Update
    Dat_Interim.Recordset.MoveNext
  Wend
  
  Set Colonnes_Select.Grd_select.DataSource = RsSelect

    Exit Sub
    
Err_InitRsSelect:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Sub ModifierDoc()

Dim myword As Object
Dim NomFichier As String
Dim CheminFichier As String
Dim fd As Field
Dim IdVue As Integer
Dim trouve As Boolean

  MousePointer = vbHourglass
  CreerRequeteDoc
  CreationFichierData
  NomFichier = Dat_wdoc.Recordset("wdoc_nom")
  If Right(CheminWord, 1) = "\" Then
    CheminFichier = CheminWord + NomFichier + ".doc"
  Else
    CheminFichier = CheminWord + "\" + NomFichier + ".doc"
  End If
  Set myword = GetObject(CheminFichier)
             
  myword.Parent.WindowState = wdWindowStateMaximize
  myword.Parent.DOCUMENTS(1).Activate
  myword.Parent.DOCUMENTS(1).Windows.Add
  myword.Parent.RecentFiles.Add Document:=myword.Parent.ActiveDocument.Name
  MousePointer = vbDefault
  myword.Parent.Visible = True
  Set myword = Nothing
    
End Sub

Sub ModifRsSelect()

Dim trouve As Boolean

  'vérifie s'il existe une base locale pour les conditions de sélection de cette source de données
  DbAccessPath = "c:\Program Files\" + App.EXEName + "\" + CStr(Dat_wdoc.Recordset("wsou_id")) + ".mdb"
  If Dir$(DbAccessPath, vbNormal) <> "" Then  'récupère les conditions de sélection
    Colonnes_Select.stb_ChoixColonnes.Panels(3).Text = "Base locale"
    Workspaces(0).OpenDatabase (DbAccessPath) 'situées dans la base locale
    With Dat_Interim
      .DatabaseName = DbAccessPath
      .RecordSource = "select * from wwhe where wsou_id=" + CStr(Dat_wdoc.Recordset("wsou_id"))
      .Refresh
    End With
    If Not Dat_Interim.Recordset.EOF Then Dat_Interim.Recordset.MoveFirst
    While Not Dat_Interim.Recordset.EOF
      'met à jour le grid
      RsSelect.MoveFirst
      trouve = False
      While Not RsSelect.EOF And Not trouve
        If Dat_Interim.Recordset("wcol_nom") = RsSelect(0) Then
          trouve = True
        Else
          RsSelect.MoveNext
        End If
      Wend
      If trouve Then
        If Dat_Interim.Recordset("wwhe_cond") <> "" Then RsSelect(2) = Dat_Interim.Recordset("wwhe_cond")
        If Dat_Interim.Recordset("wwhe_inf") <> "" Then RsSelect(3) = Dat_Interim.Recordset("wwhe_inf")
        If Dat_Interim.Recordset("wwhe_sup") <> "" Then RsSelect(4) = Dat_Interim.Recordset("wwhe_sup")
      End If
      Dat_Interim.Recordset.MoveNext
    Wend
  Else  'récupére les conditions de sélection de la base centrale
    Colonnes_Select.stb_ChoixColonnes.Panels(3).Text = "Base centrale"
    With Dat_Interim
      .DatabaseName = CheminBase
      .RecordSource = "select * from wwhe where wsou_id=" + CStr(Dat_wdoc.Recordset("wsou_id"))
      .Refresh
    End With
    While Not Dat_Interim.Recordset.EOF
      'met à jour le grid
      RsSelect.MoveFirst
      trouve = False
      While Not RsSelect.EOF And Not trouve
        If Dat_Interim.Recordset("wcol_nom") = RsSelect(0) Then
          trouve = True
        Else
          RsSelect.MoveNext
        End If
      Wend
      If trouve Then
        If Dat_Interim.Recordset("wwhe_cond") <> "" Then RsSelect(2) = Dat_Interim.Recordset("wwhe_cond")
        If Dat_Interim.Recordset("wwhe_inf") <> "" Then RsSelect(3) = Dat_Interim.Recordset("wwhe_inf")
        If Dat_Interim.Recordset("wwhe_sup") <> "" Then RsSelect(4) = Dat_Interim.Recordset("wwhe_sup")
      End If
      Dat_Interim.Recordset.MoveNext
    Wend
  End If

End Sub

Sub RemplissageTableWRES()

Dim i, j As Integer
Dim requete As String

  Dat_Interim.Recordset.MoveFirst
  While Not Dat_Interim.Recordset.EOF
    requete = "insert into wres values("
    For j = 0 To Dat_Interim.Recordset.Fields.Count - 1
    
      Select Case Dat_Interim.Recordset.Fields(j).Type
        Case dbBigInt, dbByte, dbCurrency, dbDecimal, dbDouble, dbFloat, dbInteger, dbLong, dbNumeric, dbSingle
          If IsNull(Dat_Interim.Recordset.Fields(j)) Then
            requete = requete + "NULL,"
          Else
            requete = requete + Iifent(Dat_Interim.Recordset.Fields(j).Value) + ","
          End If
        Case dbDate
          If IsNull(Dat_Interim.Recordset.Fields(j).Value) Then
            requete = requete + "NULL" + ","
          Else
            requete = requete + Iifch(Dat_Interim.Recordset.Fields(j).Value) + ","
          End If
          
        Case dbChar, dbText
          If IsNull(Dat_Interim.Recordset.Fields(j).Value) Then
            requete = requete + "NULL" + ","
          Else
            requete = requete + Iifch(Dat_Interim.Recordset.Fields(j).Value) + ","
          End If
          
        Case dbBoolean
        
      End Select
      
    Next j
    'enlève la virgule en trop
    requete = Left(requete, Len(requete) - 1)
    requete = requete + ")"
    DbAccessRes.Execute requete, dbFailOnError
    Dat_Interim.Recordset.MoveNext
  Wend

End Sub


Sub CreerDoc()

Dim myword As Object
Dim NomFichier As String
Dim CheminFichier As String
Dim fd As Field
Dim IdVue As Integer
Dim trouve As Boolean

  MousePointer = vbHourglass
  CreerRequeteDoc
  CreationFichierData
  
  NomFichier = Dat_wdoc.Recordset("wdoc_nom")
  
  If Right(CheminWord, 1) = "\" Then
    CheminFichier = CheminWord + NomFichier + ".doc"
  Else
    CheminFichier = CheminWord + "\" + NomFichier + ".doc"
  End If
  Open CheminFichier For Random As #1
  Close #1
  
  Set myword = GetObject(CheminFichier)
  myword.Parent.WindowState = wdWindowStateMaximize
  myword.Parent.DOCUMENTS(1).Activate
  myword.Parent.DOCUMENTS(1).Windows.Add
  
  With myword.Parent.ActiveDocument.MailMerge
    .MainDocumentType = wdFormLetters
    .OpenDataSource Name:=FichierDonnées, _
                    ReadOnly:=True, _
                    linktosource:=True
  End With
  myword.Parent.ActiveDocument.SaveAs FileName:=CheminFichier, fileformat:=wdFormatDocument
  MousePointer = vbDefault
  myword.Parent.Visible = True
  Set myword = Nothing
  
End Sub


Sub CreerRequete(IdVue As Integer)
Const pkstrNomProcédure As String = "CreerRequete"
On Error GoTo Err_CreerRequete

Dim sql, requete As String
Dim Idsource, Cond, Binf, Bsup As String
Dim TypeCol, NomCol, Sens As String
Dim IdCol As Integer
Dim Jour As String
Dim Mois As String
Dim An As String
Dim Dat As Date
Dim Dates As String

  'on crée la requête résultat de la sélection
  sql = "select * from wsel where wsou_id=" + CStr(Dat_wdoc.Recordset("wsou_id"))
  With Dat_Interim
    .DatabaseName = CheminBase
    .RecordSource = sql
    .Refresh
    .Recordset.MoveFirst
  End With
  requete = "select "
 
  'colonnes
  While Not Dat_Interim.Recordset.EOF
    requete = requete + Dat_Interim.Recordset("wcol_nom") + ","
    Dat_Interim.Recordset.MoveNext
  Wend
  'enlève la dernière virgule
  requete = Left(requete, Len(requete) - 1)
  requete = requete + " from " + Dat_wdoc.Recordset("wsou_vue")
    
  'conditions
  sql = "select * from wwhe where wsou_id=" + CStr(Dat_wdoc.Recordset("wsou_id"))
  Select Case Colonnes_Select.stb_ChoixColonnes.Panels(3).Text
    Case "Base centrale"
      Dat_Interim.DatabaseName = CheminBase
      
    Case "Base locale"
      Dat_Interim.DatabaseName = "c:\program files\" + App.EXEName + "\" + CStr(Dat_wdoc.Recordset("wsou_id")) + ".mdb"
  End Select
    
  With Dat_Interim
    .RecordSource = sql
    .Refresh
  End With
  If Not Dat_Interim.Recordset.EOF Then
    requete = requete + " where "
    Dat_Interim.Recordset.MoveFirst
    While Not Dat_Interim.Recordset.EOF
      NomCol = Dat_Interim.Recordset("wcol_nom")
      If Not IsNull(Dat_Interim.Recordset("wwhe_cond")) Then
        'recherche identifiant de colonne
        IdCol = 0
        While Db.QueryDefs(IdVue).Fields(IdCol).Name <> NomCol And IdCol < Db.QueryDefs(IdVue).Fields.Count - 1
          IdCol = IdCol + 1
        Wend
        TypeCol = Db.QueryDefs(IdVue).Fields(IdCol).Type
        Select Case TypeCol
          Case dbBigInt, dbByte, dbCurrency, dbDecimal, dbDouble, dbFloat, dbInteger, dbLong, dbNumeric, dbSingle
            requete = requete + NomCol + "=" + CStr(Dat_Interim.Recordset("wwhe_cond")) + " and "
            
          Case dbDate
            Dat = CDate(Dat_Interim.Recordset.Fields("wwhe_cond"))
            Jour = Day(Dat)
            Mois = Month(Dat)
            An = Year(Dat)
            Dates = Mois + "/" + Jour + "/" + An
            requete = requete + NomCol + "=#" + Dates + "# and "
            
          Case dbChar, dbText
            requete = requete + NomCol + "=" + Iifch(Dat_Interim.Recordset("wwhe_cond")) + " and "
            
          Case dbBoolean
            requete = requete & NomCol & "=" & Dat_Interim.Recordset("wwhe_cond") + " and "
          
        End Select
      End If
      If Not IsNull(Dat_Interim.Recordset("wwhe_inf")) Then
        'recherche identifiant de colonne
        IdCol = 0
        While Db.QueryDefs(IdVue).Fields(IdCol).Name <> NomCol And IdCol < Db.QueryDefs(IdVue).Fields.Count - 1
          IdCol = IdCol + 1
        Wend
        TypeCol = Db.QueryDefs(IdVue).Fields(IdCol).Type
        
        Select Case TypeCol
          Case dbBigInt, dbByte, dbCurrency, dbDecimal, dbDouble, dbFloat, dbInteger, dbLong, dbNumeric, dbSingle
            requete = requete + NomCol + ">" + CStr(Dat_Interim.Recordset("wwhe_inf")) + " and "
            
          Case dbDate
            Dat = CDate(Dat_Interim.Recordset.Fields("wwhe_inf"))
            Jour = Day(Dat)
            Mois = Month(Dat)
            An = Year(Dat)
            Dates = Mois + "/" + Jour + "/" + An
            requete = requete + NomCol + ">#" + Dates + "# And "
            
          Case dbChar, dbText
            requete = requete + NomCol + ">" + Iifch(Dat_Interim.Recordset("wwhe_inf")) + " and "
            
          Case dbBoolean
          
        End Select
        
      End If
      If Not IsNull(Dat_Interim.Recordset("wwhe_sup")) Then
        'recherche identifiant de colonne
        IdCol = 0
        While Db.QueryDefs(IdVue).Fields(IdCol).Name <> NomCol And IdCol < Db.QueryDefs(IdVue).Fields.Count - 1
          IdCol = IdCol + 1
        Wend
        TypeCol = Db.QueryDefs(IdVue).Fields(IdCol).Type
        
        Select Case TypeCol
          Case dbBigInt, dbByte, dbCurrency, dbDecimal, dbDouble, dbFloat, dbInteger, dbLong, dbNumeric, dbSingle
            requete = requete + NomCol + "<" + CStr(Dat_Interim.Recordset("wwhe_sup")) + " and "
            
          Case dbDate
            Dat = CDate(Dat_Interim.Recordset.Fields("wwhe_sup"))
            Jour = Day(Dat)
            Mois = Month(Dat)
            An = Year(Dat)
            Dates = Mois + "/" + Jour + "/" + An
            requete = requete + NomCol + "<#" + Dates + "# and "
            
          Case dbChar, dbText
            requete = requete + NomCol + "<" + Iifch(Dat_Interim.Recordset("wwhe_sup")) + " and "
            
          Case dbBoolean

        End Select
      End If
  
       'passe à l'enregistrement suivant
      Dat_Interim.Recordset.MoveNext
    Wend
    'enlève le " and " en trop
    requete = Left(requete, Len(requete) - 5)
  End If
   
  'tri
  sql = "select * from woby where wsou_id=" + CStr(Dat_wdoc.Recordset("wsou_id")) + " order by woby_numordre"
  Dat_Interim.DatabaseName = CheminBase
  Dat_Interim.RecordSource = sql
  Dat_Interim.Refresh
  If Not Dat_Interim.Recordset.EOF Then
    requete = requete + " order by "
    Dat_Interim.Recordset.MoveFirst
    While Not Dat_Interim.Recordset.EOF
      Select Case Dat_Interim.Recordset.Fields("woby_sens")
        Case "A"
          Sens = " asc,"
        Case "D"
          Sens = " desc,"
      End Select
      requete = requete + Dat_Interim.Recordset.Fields("wcol_nom") + Sens
      Dat_Interim.Recordset.MoveNext
    Wend
  End If
  'enleve la virgule en trop
  If Right(requete, 1) = "," Then requete = Left(requete, Len(requete) - 1)
  
  Dat_Interim.RecordSource = requete
  Dat_Interim.Refresh

    Exit Sub
    
Err_CreerRequete:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


Sub RempliTableWresUnEnreg()

Dim i, j As Integer
Dim requete As String

  Dat_Interim.Recordset.MoveFirst
    requete = "insert into wres values("
    For j = 0 To Dat_Interim.Recordset.Fields.Count - 1
    
      Select Case Dat_Interim.Recordset.Fields(j).Type
        Case dbBigInt, dbByte, dbCurrency, dbDecimal, dbDouble, dbFloat, dbInteger, dbLong, dbNumeric, dbSingle
          If IsNull(Dat_Interim.Recordset.Fields(j)) Then
            requete = requete + "NULL,"
          Else
            requete = requete + Iifent(Dat_Interim.Recordset.Fields(j).Value) + ","
          End If
        Case dbDate
          If IsNull(Dat_Interim.Recordset.Fields(j).Value) Then
            requete = requete + "NULL" + ","
          Else
            requete = requete + Iifch(Dat_Interim.Recordset.Fields(j).Value) + ","
          End If
          
        Case dbChar, dbText
          If IsNull(Dat_Interim.Recordset.Fields(j).Value) Then
            requete = requete + "NULL" + ","
          Else
            requete = requete + Iifch(Dat_Interim.Recordset.Fields(j).Value) + ","
          End If
          
        Case dbBoolean
        
      End Select
      
    Next j
    'enlève la virgule en trop
    requete = Left(requete, Len(requete) - 1)
    requete = requete + ")"
    DbAccessRes.Execute requete, dbFailOnError

End Sub

Private Sub Cmd_Actualiser_Click()

  Dim sel_famille, sel_source, sel_docident, sel_docnom, sql As String
  
  Me.MousePointer = vbHourglass
  If Txt_Famille <> "" Then sel_famille = " AND WFAM_NOM LIKE """ & Txt_Famille & "*"""
  If Txt_Source <> "" Then sel_source = " AND WSOU_NOM LIKE """ & Txt_Source & "*"""
  If Txt_DocIdent <> "" Then sel_docident = " AND WDOC_ID LIKE """ & Txt_DocIdent & "*"""
  If Txt_DocNom <> "" Then sel_docnom = " AND WDOC_NOM LIKE """ & Txt_DocNom & "*"""
  
  sql = "select wdoc.*,wfam.wfam_nom,wsou.wsou_nom,wsou.wsou_vue from wdoc,wfam,wsou "
  sql = sql + "where wdoc.wfam_id=wfam.wfam_id and wdoc.wsou_id=wsou.wsou_id"
  sql = sql + sel_famille + sel_source + sel_docident + sel_docnom
  sql = sql + " order by wdoc_id"
  
  Dat_wdoc.RecordSource = sql
  Dat_wdoc.Refresh
  Me.MousePointer = vbDefault
  
End Sub

Private Sub cmd_Creer_Click()

  stb_ListeDoc.Panels(4).Text = "CREATION"
  Load DocumentWord_fiche
  DocumentWord_fiche.Caption = "Création d'un document Word"

  If Dat_wdoc.Recordset.RecordCount = 0 Then
    DocumentWord_fiche.Txt_Ident = 1
  Else
    With Dat_wdoc
      .Recordset.MoveFirst
      .Recordset.MoveLast
    End With
    DocumentWord_fiche.Dat_Id.RecordSource = "select max(wdoc_id) from wdoc"
    DocumentWord_fiche.Dat_Id.Refresh
    DocumentWord_fiche.Txt_Ident.Text = DocumentWord_fiche.Dat_Id.Recordset(0) + 1
    DocumentWord_fiche.Txt_Nom = ""
  End If
  DocumentWord_fiche.Show vbModal
  stb_ListeDoc.Panels(3).Text = ""

End Sub

Private Sub Cmd_DocIdent_Click()

  Txt_DocIdent.Text = ""
  Txt_DocIdent.SetFocus
  
End Sub

Private Sub Cmd_DocNom_Click()

  Txt_DocNom.Text = ""
  Txt_DocNom.SetFocus
  
End Sub


Private Sub Cmd_Famille_Click()

  Txt_Famille.Text = ""
  Txt_Famille.SetFocus
  
End Sub

Private Sub Cmd_Mailings_Click()

Dim sql As String
Dim vntTemp As Variant

  If Not Dat_wdoc.Recordset.EOF Then
    sql = "select * from wsel where wsou_id=" + CStr(Dat_wdoc.Recordset("wsou_id"))
    Dat_Interim.DatabaseName = CheminBase
    Dat_Interim.RecordSource = sql
    Dat_Interim.Refresh
    If Dat_Interim.Recordset.EOF Then
      MsgBox "Vous ne pouvez pas éditer ce mailing." + Chr(13) + _
             "Sa source de données ne contient pas de colonnes !", vbExclamation, "ERREUR"
    Else
      Load Colonnes_Select
      
      Colonnes_Select.Vue = Dat_wdoc.Recordset("wsou_vue")
      Colonnes_Select.FeuilleAppelante = Me.Name
      Colonnes_Select.stb_ChoixColonnes.Panels(4).Text = "FUSION"
      Colonnes_Select.Cmd_Defaut.Visible = True
      Colonnes_Select.cmd_Modifier.Visible = True
      Colonnes_Select.Grd_select.Columns(4).Locked = True
      Colonnes_Select.Grd_select.Columns(3).Locked = True
      Colonnes_Select.Grd_select.Columns(2).Locked = True
      Colonnes_Select.Grd_select.Columns(1).Locked = True
      Colonnes_Select.Grd_select.Columns(0).Locked = True
      InitRsSelect
      ModifRsSelect
      
      Colonnes_Select.Show vbModal
    End If
  End If
    
End Sub

Private Sub cmd_Modifier_Click()

  If Not Dat_wdoc.Recordset.EOF Then
    If Dat_wdoc.Recordset("wdoc_bloque") = 1 Then
      MsgBox "Vous ne pouvez pas modifier ce document.", vbExclamation, "IMPOSSIBLE"
    Else
      stb_ListeDoc.Panels(4).Text = "MODIFICATION"
      Load DocumentWord_fiche
      With DocumentWord_fiche
        .Caption = "Modification d'un document Word"
        .Txt_Ident = Dat_wdoc.Recordset("wdoc_id")
        .Txt_Nom = Dat_wdoc.Recordset("wdoc_nom")
        .txt_IdFam = Dat_wdoc.Recordset("wfam_id")
        .Txt_Famille = Dat_wdoc.Recordset("wfam_nom")
        .Txt_IdSource = Dat_wdoc.Recordset("wsou_id")
        .Txt_Source = Dat_wdoc.Recordset("wsou_nom")
      End With
      DocumentWord_fiche.Cmb_Bloque.ListIndex = 1 'non
      DocumentWord_fiche.Cmd_Source.Enabled = True
      DocumentWord_fiche.Txt_Nom.Enabled = False
      DocumentWord_fiche.Cmd_ModifierWord.Visible = True
      DocumentWord_fiche.Show vbModal
      stb_ListeDoc.Panels(3).Text = ""
    End If
  End If

End Sub


Private Sub cmd_Quitter_Click()
  Unload Me
End Sub

Private Sub Cmd_Source_Click()

  Txt_Source.Text = ""
  Txt_Source.SetFocus
  
End Sub

Private Sub cmd_Supprimer_Click()

  If Not Dat_wdoc.Recordset.EOF Then
    If Dat_wdoc.Recordset("wdoc_bloque") = 1 Then
      MsgBox "Vous ne pouvez pas supprimer ce document.", vbExclamation, "IMPOSSIBLE"
    Else
      stb_ListeDoc.Panels(4).Text = "SUPPRESSION"
      Load DocumentWord_fiche
      With DocumentWord_fiche
        .Caption = "Suppression d'un document Word"
        .Txt_Ident = Dat_wdoc.Recordset("wdoc_id")
        .Txt_Nom = Dat_wdoc.Recordset("wdoc_nom")
        .Txt_Nom.Enabled = False
        .txt_IdFam = Dat_wdoc.Recordset("wfam_id")
        .Txt_Famille = Dat_wdoc.Recordset("wfam_nom")
        .Txt_IdSource = Dat_wdoc.Recordset("wsou_id")
        .Txt_Source = Dat_wdoc.Recordset("wsou_nom")
        .Cmd_Famille.Enabled = False
        .Cmd_Source.Enabled = False
        .Cmb_Bloque.ListIndex = 1 'non
        .Cmb_Bloque.Enabled = False
      End With
      
      DocumentWord_fiche.Show vbModal
      stb_ListeDoc.Panels(3).Text = ""
    End If
  End If
  
End Sub







Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)

  Dim colonne As String
  Dim classem As String
  
  colonne = DBGrid1.Columns(ColIndex).DataField
  Call Tri(Dat_wdoc, DBGrid1, colonne)
      
  
  'On affiche dans la statusbar l'ordre de tri
  If InStr(Dat_wdoc.RecordSource, "ASC") <> 0 Then
      classem = "Croissant"
  Else
      classem = "Décroissant"
  End If
         
  stb_ListeDoc.Panels(4).Text = "Classement " & classem

End Sub


Private Sub Form_Load()

Dim sql As String

  Dat_Interim.DatabaseName = CheminBase
  
  sql = "select wdoc.*,wfam.wfam_nom,wsou.wsou_nom,wsou.wsou_vue from wdoc,wfam,wsou "
  sql = sql + "where wdoc.wfam_id=wfam.wfam_id and wdoc.wsou_id=wsou.wsou_id"
  sql = sql + " order by wdoc_id"
  
  With Dat_wdoc
    .DatabaseName = CheminBase
    .RecordSource = sql
    .Refresh
  End With
  
End Sub



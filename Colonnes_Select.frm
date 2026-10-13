VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Colonnes_Select 
   Caption         =   "Conditions de sélection des enregistrements"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   ControlBox      =   0   'False
   Icon            =   "Colonnes_Select.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8775
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid Grd_select 
      Height          =   3855
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      TabAction       =   2
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Nomcolonne"
         Caption         =   "Nom colonne"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "alias"
         Caption         =   "Nom de la colonne"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "condition"
         Caption         =   "Egal à ou commence par"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "borneinf"
         Caption         =   "Borne inférieure"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "bornesup"
         Caption         =   "Borne supérieure"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   2700,284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2055,118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1409,953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1484,787
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.CommandButton cmd_Modifier 
         Height          =   375
         Left            =   1080
         Picture         =   "Colonnes_Select.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Modifier le document Word sélectionné"
         Top             =   150
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Cmd_Defaut 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   360
         Picture         =   "Colonnes_Select.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Récuperer les conditions de sélection par défaut"
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
         Height          =   345
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         ToolTipText     =   "Récuperer les critères de sélection par défaut"
         Top             =   120
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   7800
         Picture         =   "Colonnes_Select.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Annuler 
         Height          =   375
         Left            =   7080
         Picture         =   "Colonnes_Select.frx":0748
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_ChoixColonnes 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4950
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
            TextSave        =   "25/07/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "14:49"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Colonnes_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Vue As String    'Ces 2 variables sont utilisées par la fonction VerifType
Public FeuilleAppelante As String


Sub CreationBaseIdsource(Idsource As String)
Const pkstrNomProcédure As String = "CreationBaseIdsource"
On Error GoTo Err_CreationBaseIdsource

'Création de la base de données ayant pour nom : Nom de l'exe + Identifiant de la source de données
'si elle n'existe pas déjà
Dim td As TableDef
Dim fd As Field
Dim id As Index

  DbAccessPath = "c:\Program Files\" + App.EXEName + "\" + Idsource + ".mdb"
  If Dir$(DbAccessPath, vbNormal) = "" Then
    'création de la base
    Set DbAccess = Workspaces(0).CreateDatabase(DbAccessPath, dbLangGeneral, dbVersion30)

    'création de la table WWHE
    Set td = DbAccess.CreateTableDef("WWHE")
    'création des champs de la table
    Set fd = td.CreateField("WSOU_ID", dbInteger)
    'ajout du champ à la collection des champs de la table
    td.Fields.Append fd
    Set fd = td.CreateField("WCOL_NOM", dbText, daoWCOL_NOM)
    td.Fields.Append fd
    Set fd = td.CreateField("WWHE_COND", dbText, 30)
    td.Fields.Append fd
    Set fd = td.CreateField("WWHE_INF", dbText, 30)
    td.Fields.Append fd
    Set fd = td.CreateField("WWHE_SUP", dbText, 30)
    td.Fields.Append fd
  
    'ajout de la table à la collection des tables
    DbAccess.TableDefs.Append td
  Else
    Set DbAccess = Workspaces(0).OpenDatabase(DbAccessPath)
  End If
  
    Exit Sub
    
Err_CreationBaseIdsource:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Sub CreationTableWRES()
'création de la table dans la nouvelle base

Dim td As TableDef
Dim fd As Field
Dim id As Index
Dim i As Integer
Dim NomCol As String

  'création de la table
  Set td = DbAccess.CreateTableDef("WRES")
  'création des champs de la table
  For i = 0 To Dat_Interim.Recordset.Fields.Count - 1
    NomCol = Dat_Interim.Recordset.Fields(i).Name
    Set fd = td.CreateField(NomCol, Dat_Interim.Recordset.Fields(i).Type, Dat_Interim.Recordset.Fields(i).Size)
    'ajout du champ à la collection des champs de la table
    td.Fields.Append fd
  Next i

  'ajout de la table à la collection des tables
  DbAccess.TableDefs.Append td
  
End Sub


Sub RemplissageTableWRES()

Dim i, j As Integer
Dim requete As String

  Dat_Interim.Recordset.MoveFirst
    
  'construction de la requete
  While Not Dat_Interim.Recordset.EOF
    requete = "insert into wres values("
    For j = 0 To Dat_Interim.Recordset.Fields.Count - 1
      Select Case Dat_Interim.Recordset.Fields(j).Type
        Case 3, 4 'entier
          requete = requete + Iifent(Dat_Interim.Recordset.Fields(j).Value) + ","
        Case 8  'date
          requete = requete + Iifch(Dat_Interim.Recordset.Fields(j).Value) + ","
        Case 10 'text
          requete = requete + Iifch(Dat_Interim.Recordset.Fields(j).Value) + ","
      End Select
    Next j
    'enlève la virgule en trop
    requete = Left(requete, Len(requete) - 1)
    requete = requete + ")"
    DbAccess.Execute requete, dbFailOnError
  'Next i
    Dat_Interim.Recordset.MoveNext
  Wend

End Sub


Function VerifType() As Boolean

  Dim sql, NomCol, TypeCol As String
  Dim IdVue, IdCol, Valeur, i As Integer
  Dim trouve As Boolean
  
  'recherche l'identifiant de la vue
  trouve = False
  IdVue = 0
  While Not trouve And IdVue <= Db.QueryDefs.Count - 1
    If Db.QueryDefs(IdVue).Name = Vue Then
      trouve = True
    Else
      IdVue = IdVue + 1
    End If
  Wend
  
  Select Case FeuilleAppelante
    Case "SourcesDonnées_fiche"
      If Not trouve Then
        MsgBox "La vue utilisée par la source de données n'existe pas.", vbExclamation, "ERREUR"
        VerifType = False
        Exit Function
      Else
        SourcesDonnées_fiche.RsSelect.MoveFirst
        While Not SourcesDonnées_fiche.RsSelect.EOF
          'recherche l'identifiant de la colonne
          IdCol = 0
          trouve = False
          NomCol = SourcesDonnées_fiche.RsSelect(0)
          While Not trouve And IdCol <= Db.QueryDefs(IdVue).Fields.Count - 1
            If Db.QueryDefs(IdVue).Fields(IdCol).Name = NomCol Then
              trouve = True
            Else
              IdCol = IdCol + 1
            End If
          Wend
          If Not trouve Then
            MsgBox "Le champ " + NomCol + " n'existe pas dans la vue " + Db.QueryDefs(IdVue).Name + ".", vbExclamation, "ERREUR"
            VerifType = False
            Exit Function
          End If
          'type de la colonne
          TypeCol = Db.QueryDefs(IdVue).Fields(IdCol).Type
          
          Select Case TypeCol
            Case dbBigInt, dbByte, dbCurrency, dbDecimal, dbDouble, dbFloat, dbInteger, dbLong, dbNumeric, dbSingle
              For i = 2 To 4
                If Not IsNull(SourcesDonnées_fiche.RsSelect(i)) And Not IsNumeric(SourcesDonnées_fiche.RsSelect(i)) Then
                    MsgBox "Le champ " + SourcesDonnées_fiche.RsSelect(1) + " est de type numérique" + Chr(13) + _
                           "Veuillez corriger votre condition, ce doit être un numérique."
                    VerifType = False
                    Exit Function
                End If
              Next i
              VerifType = True
            
            Case dbBoolean
              For i = 2 To 4
                If Not IsNull(SourcesDonnées_fiche.RsSelect(i)) And SourcesDonnées_fiche.RsSelect(i).Value <> "TRUE" And SourcesDonnées_fiche.RsSelect(i).Value <> "FALSE" Then
                  MsgBox "Le champ " + SourcesDonnées_fiche.RsSelect(1) + " est de type booléen ." + Chr(13) + _
                         "Veuillez corriger votre condition, elle doit être égale à TRUE ou FALSE."
                  VerifType = False
                  Exit Function
                End If
              Next i
              VerifType = True
            
            Case dbChar, dbText 'pas besoin de verif tout peut être une chaine
              VerifType = True
                        
            Case dbDate
              For i = 2 To 4
                If Not IsNull(SourcesDonnées_fiche.RsSelect(i)) And Not IsDate(SourcesDonnées_fiche.RsSelect(i)) Then
                  MsgBox "Le champ " + SourcesDonnées_fiche.RsSelect(1) + " est de type date ." + Chr(13) + _
                         "Veuillez corriger votre condition, ce doit être une date."
                  VerifType = False
                  Exit Function
                End If
              Next i
              VerifType = True
              
            Case Else
              VerifType = True
            
          End Select
          SourcesDonnées_fiche.RsSelect.MoveNext
        Wend
      End If

    Case "DocumentWord_liste"
      If Not trouve Then
        MsgBox "La vue utilisée par la source de données n'existe pas.", vbExclamation, "ERREUR"
        VerifType = False
        Exit Function
      Else
        DocumentWord_liste.RsSelect.MoveFirst
        While Not DocumentWord_liste.RsSelect.EOF
          'recherche l'identifiant de la colonne
          IdCol = 0
          trouve = False
          NomCol = DocumentWord_liste.RsSelect(0)
          While Not trouve And IdCol <= Db.QueryDefs(IdVue).Fields.Count - 1
            If Db.QueryDefs(IdVue).Fields(IdCol).Name = NomCol Then
              trouve = True
            Else
              IdCol = IdCol + 1
            End If
          Wend
          If Not trouve Then
            MsgBox "Le champ " + NomCol + " n'existe pas dans la vue " + Db.QueryDefs(IdVue).Name + ".", vbExclamation, "ERREUR"
            VerifType = False
            Exit Function
          End If
          'type de la colonne
          TypeCol = Db.QueryDefs(IdVue).Fields(IdCol).Type
          
          Select Case TypeCol
            Case dbBigInt, dbByte, dbCurrency, dbDecimal, dbDouble, dbFloat, dbInteger, dbLong, dbNumeric, dbSingle
              For i = 2 To 4
                If Not IsNull(DocumentWord_liste.RsSelect(i)) And Not IsNumeric(DocumentWord_liste.RsSelect(2)) Then
                    MsgBox "Le champ " + DocumentWord_liste.RsSelect(1) + " est de type numérique" + Chr(13) + _
                           "Veuillez corriger votre condition, ce doit être un numérique."
                    VerifType = False
                    Exit Function
                End If
              Next i
              VerifType = True
            
            Case dbBoolean
              For i = 2 To 4
                If Not IsNull(DocumentWord_liste.RsSelect(i)) And DocumentWord_liste.RsSelect(i).Value <> "TRUE" And DocumentWord_liste.RsSelect(i).Value <> "FALSE" Then
                  MsgBox "Le champ " + DocumentWord_liste.RsSelect(1) + " est de type booléen ." + Chr(13) + _
                         "Veuillez corriger votre condition, elle doit être égale à TRUE ou FALSE."
                  VerifType = False
                  Exit Function
                End If
              Next i
              VerifType = True
            
            Case dbChar, dbText 'pas besoin de verif tout peut être une chaine
              VerifType = True
                        
            Case dbDate
              For i = 2 To 4
                If Not IsNull(DocumentWord_liste.RsSelect(i)) And Not IsDate(DocumentWord_liste.RsSelect(i)) Then
                  MsgBox "Le champ " + DocumentWord_liste.RsSelect(1) + " est de type date ." + Chr(13) + _
                         "Veuillez corriger votre condition, ce doit être une date."
                  VerifType = False
                  Exit Function
                End If
              Next i
              VerifType = True
              
            Case Else
              VerifType = True
            
          End Select
          DocumentWord_liste.RsSelect.MoveNext
        Wend
      End If
    
  End Select
    
End Function


Private Sub cmd_Annuler_Click()
  Unload Me
End Sub


Private Sub Cmd_Defaut_Click()

Dim sql As String
  
    sql = "select * from wsel where wsou_id=" + CStr(DocumentWord_liste.Dat_wdoc.Recordset("wsou_id"))
    DocumentWord_liste.Dat_Interim.RecordSource = sql
    DocumentWord_liste.Dat_Interim.DatabaseName = CheminBase
    DocumentWord_liste.Dat_Interim.Refresh
    DocumentWord_liste.InitRsSelect
    DocumentWord_liste.DefautRsSelect
    Grd_select.Columns(4).Locked = True
    Grd_select.Columns(3).Locked = True
    Grd_select.Columns(2).Locked = True
    Grd_select.Columns(1).Locked = True
    Grd_select.Columns(0).Locked = True
    stb_ChoixColonnes.Panels(3).Text = "Base centrale"
    
End Sub

Private Sub cmd_Modifier_Click()


  Grd_select.Columns(4).Locked = False
  Grd_select.Columns(3).Locked = False
  Grd_select.Columns(2).Locked = False
  Grd_select.Columns(1).Locked = False
  Grd_select.Columns(0).Locked = False

  stb_ChoixColonnes.Panels(3).Text = "Base locale"
  
End Sub

Private Sub cmd_Valider_Click()

Dim sql As String
Dim Cond, Binf, Bsup As String

On Error GoTo Rupture_Integrite

  Select Case stb_ChoixColonnes.Panels(4).Text
    Case "CREATION"
      MousePointer = vbHourglass
      'vérifie si le type des conditions de sélection saisies correspond à celui des champs
      If Not VerifType Then
        MousePointer = vbDefault
        Exit Sub
      Else
        'crée les enregistrements dans WWHE dans la base partagée sur le serveur
        SourcesDonnées_fiche.RsSelect.MoveFirst
        While Not SourcesDonnées_fiche.RsSelect.EOF
        
          If Not IsNull(SourcesDonnées_fiche.RsSelect.Fields(2)) Then
            Cond = Iifch(SourcesDonnées_fiche.RsSelect.Fields(2))
          Else
            Cond = "NULL"
          End If
          If Not IsNull(SourcesDonnées_fiche.RsSelect.Fields(3)) Then
            Binf = Iifch(SourcesDonnées_fiche.RsSelect.Fields(3))
          Else
            Binf = "NULL"
          End If
          If Not IsNull(SourcesDonnées_fiche.RsSelect.Fields(4)) Then
            Bsup = Iifch(SourcesDonnées_fiche.RsSelect.Fields(4))
          Else
            Bsup = "NULL"
          End If
          
          If Cond <> "NULL" Or Binf <> "NULL" Or Bsup <> "NULL" Then
            sql = "insert into wwhe values(" + CStr(SourcesDonnées_fiche.Txt_id) + ","
            sql = sql + Iifch(SourcesDonnées_fiche.RsSelect.Fields(0)) + "," 'Cette colonne est invisible
            sql = sql + Cond + "," + Binf + "," + Bsup + ")"
            Db.Execute sql, dbFailOnError
          End If
          SourcesDonnées_fiche.RsSelect.MoveNext
        Wend
        SourcesDonnées_fiche.SelectFait = True
      End If
    
    Case "MODIFICATION"
      MousePointer = vbHourglass
      'vérifie si le type des conditions de sélection saisies correspond à celui des champs
      If Not VerifType Then
        MousePointer = vbDefault
        Exit Sub
      Else
        sql = "delete * from wwhe where wsou_id=" + CStr(SourcesDonnées_fiche.Txt_id)
        Db.Execute sql, dbFailOnError
        SourcesDonnées_fiche.RsSelect.MoveFirst
        While Not SourcesDonnées_fiche.RsSelect.EOF
        
          If Not IsNull(SourcesDonnées_fiche.RsSelect.Fields(2)) Then
            Cond = Iifch(SourcesDonnées_fiche.RsSelect.Fields(2))
          Else
            Cond = "NULL"
          End If
          If Not IsNull(SourcesDonnées_fiche.RsSelect.Fields(3)) Then
            Binf = Iifch(SourcesDonnées_fiche.RsSelect.Fields(3))
          Else
            Binf = "NULL"
          End If
          If Not IsNull(SourcesDonnées_fiche.RsSelect.Fields(4)) Then
            Bsup = Iifch(SourcesDonnées_fiche.RsSelect.Fields(4))
          Else
            Bsup = "NULL"
          End If
    
          If Cond <> "NULL" Or Binf <> "NULL" Or Bsup <> "NULL" Then
            sql = "insert into wwhe values(" + CStr(SourcesDonnées_fiche.Txt_id) + ","
            sql = sql + Iifch(SourcesDonnées_fiche.RsSelect.Fields(0)) + "," 'Cette colonne est invisible
            sql = sql + Cond + "," + Binf + "," + Bsup + ")"
            Db.Execute sql, dbFailOnError
          End If
          SourcesDonnées_fiche.RsSelect.MoveNext
        Wend
      End If
      
    Case "FUSION"
    
      Select Case stb_ChoixColonnes.Panels(3).Text
        Case "Base locale"
          'on sauvegarde les conditions de sélection dans la base locale
          MousePointer = vbHourglass
          'vérifie si le type des conditions de sélection saisies correspond à celui des champs
          If Not VerifType Then
            MousePointer = vbDefault
            Exit Sub
          Else
            'crée une base locale IdSource.mdb (si elle n'existe pas déjà) et la table WWHE
            CreationBaseIdsource (DocumentWord_liste.Dat_wdoc.Recordset("wsou_id"))
            'vide la table WWHE
            sql = "delete * from wwhe"
            DbAccess.Execute sql, dbFailOnError
            'crée les enregistrements dans WWHE dans la base locale
            DocumentWord_liste.RsSelect.MoveFirst
            While Not DocumentWord_liste.RsSelect.EOF
              If Not IsNull(DocumentWord_liste.RsSelect.Fields(2)) Then
                Cond = Iifch(DocumentWord_liste.RsSelect.Fields(2))
              Else
                Cond = "NULL"
              End If
              If Not IsNull(DocumentWord_liste.RsSelect.Fields(3)) Then
                Binf = Iifch(DocumentWord_liste.RsSelect.Fields(3))
              Else
                Binf = "NULL"
              End If
              If Not IsNull(DocumentWord_liste.RsSelect.Fields(4)) Then
                Bsup = Iifch(DocumentWord_liste.RsSelect.Fields(4))
              Else
                Bsup = "NULL"
              End If
        
              If Cond <> "NULL" Or Binf <> "NULL" Or Bsup <> "NULL" Then
                sql = "insert into wwhe values(" + CStr(DocumentWord_liste.Dat_wdoc.Recordset("wsou_id")) + ","
                sql = sql + Iifch(DocumentWord_liste.RsSelect.Fields(0)) + "," 'Cette colonne est invisible
                sql = sql + Cond + "," + Binf + "," + Bsup + ")"
                DbAccess.Execute sql, dbFailOnError
              End If
              DocumentWord_liste.RsSelect.MoveNext
            Wend
            DocumentWord_liste.FusionDoc
          End If
        
        Case "Base centrale"
          DocumentWord_liste.FusionDoc
      End Select
      
  End Select

  MousePointer = vbDefault
  Unload Me
  
  Exit Sub
  
Rupture_Integrite:
'Affiche le code d'erreur approprié
  Select Case Err
    Case 13
        MsgBox "Valeur incompatible avec le type de colonne", vbOKOnly + vbExclamation
  
    Case 3022
        MsgBox "Doublon" + ", création impossible", vbOKOnly + vbExclamation
  
    Case 3200
        MsgBox "Référence" + ", suppression impossible", vbOKOnly + vbExclamation
  
    Case 3075
        MsgBox "L'un des champs contient une cote, mise à jour impossible", vbOKOnly + vbExclamation
  
    Case 3315
        MsgBox "Un champ obligatoire n'est pas renseigné, mise à jour impossible", vbOKOnly + vbExclamation
  
    Case Else
        MsgBox "Erreur" + CStr(Err) + " " + Err.Description + " survenue, mise à jour impossible", vbOKOnly + vbExclamation
  End Select
  MousePointer = vbDefault
  Exit Sub
  
End Sub


Private Sub Form_Load()

  Dat_Interim.DatabaseName = CheminBase
  
End Sub




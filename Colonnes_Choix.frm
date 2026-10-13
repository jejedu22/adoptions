VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Colonnes_Choix 
   Caption         =   "Choix des colonnes"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   ControlBox      =   0   'False
   Icon            =   "Colonnes_Choix.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8805
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid Grd_Colonne 
      Height          =   4050
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Liste des colonnes disponibles pour la source de données"
      Top             =   960
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   7144
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "identsource"
         Caption         =   "Id Source"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "numordre"
         Caption         =   "N° ordre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "nomcolonne"
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
      BeginProperty Column03 
         DataField       =   "aliascolonne"
         Caption         =   "Alias colonne"
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
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   810,142
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   2865,26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3240
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.CommandButton cmd_Annuler 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   7080
         Picture         =   "Colonnes_Choix.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   7800
         Picture         =   "Colonnes_Choix.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_ChoixColonnes 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5490
      Width           =   8805
      _ExtentX        =   15531
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
            TextSave        =   "14:31"
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
Attribute VB_Name = "Colonnes_Choix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmd_Annuler_Click()
  Unload Me
End Sub

Private Sub cmd_Valider_Click()

Dim bloque As Integer
Dim Alias, sql As String
On Error GoTo Rupture_Integrite

  Select Case stb_ChoixColonnes.Panels(4).Text
    Case "CREATION"
    
      SourcesDonnées_fiche.RsColonne.MoveFirst
      While Not SourcesDonnées_fiche.RsColonne.EOF
        If Not IsNull(SourcesDonnées_fiche.RsColonne.Fields(3)) Then
          Alias = Iifch(SourcesDonnées_fiche.RsColonne.Fields(3))
          'modifié le 18/05/2010 par Caroline (précision sur l'ordre des champs)
          sql = "INSERT INTO WSEL (WSOU_ID,WCOL_NOM,WSEL_NUMORDRE,WSEL_ALIAS) VALUES("
          sql = sql + CStr(SourcesDonnées_fiche.RsColonne.Fields(0)) + ","
          sql = sql + Iifch(SourcesDonnées_fiche.RsColonne.Fields(2)) + ","
          sql = sql + CStr(SourcesDonnées_fiche.RsColonne.Fields(1)) + ","
          sql = sql + Alias + ")"
          Workspaces(0).Databases(0).Execute sql, dbFailOnError
        End If
        
        SourcesDonnées_fiche.RsColonne.MoveNext
      Wend
      SourcesDonnées_fiche.ChoixFait = True
      
    Case "MODIFICATION"
      'supprime tous les enregistrements de WSEL pour la source de données modifiée
      sql = "delete * from wsel where wsou_id=" + CStr(SourcesDonnées_fiche.Txt_id)
      Workspaces(0).Databases(0).Execute sql, dbFailOnError
      'crée les enregistrements dans WSEL pour la source de données modifiée
      SourcesDonnées_fiche.RsColonne.MoveFirst
      While Not SourcesDonnées_fiche.RsColonne.EOF
        If Not IsNull(SourcesDonnées_fiche.RsColonne.Fields(3)) Then
          Alias = Iifch(SourcesDonnées_fiche.RsColonne.Fields(3))
          'modifié le 18/05/2010 par Caroline (précision sur l'ordre des champs)
          sql = "INSERT INTO WSEL (WSOU_ID,WCOL_NOM,WSEL_NUMORDRE,WSEL_ALIAS) VALUES("
          sql = sql + CStr(SourcesDonnées_fiche.RsColonne.Fields(0)) + ","
          sql = sql + Iifch(SourcesDonnées_fiche.RsColonne.Fields(2)) + ","
          sql = sql + CStr(SourcesDonnées_fiche.RsColonne.Fields(1)) + ","
          sql = sql + Alias + ")"
          Workspaces(0).Databases(0).Execute sql, dbFailOnError
        End If
        SourcesDonnées_fiche.RsColonne.MoveNext
      Wend
      'supprime les enregistrements dans WOBY pour lesquels la colonne ne fait plus partie de la sélection
      sql = "delete * from woby where wsou_id=" + CStr(SourcesDonnées_fiche.Txt_id)
      sql = sql + " and wcol_nom not in (select wcol_nom from wsel where wsou_id=" + CStr(SourcesDonnées_fiche.Txt_id) + ")"
      Workspaces(0).Databases(0).Execute sql, dbFailOnError
      
  End Select

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



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form DocumentWord_fiche 
   Caption         =   "Fiche Document Word"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   ControlBox      =   0   'False
   Icon            =   "DocumentWord_fiche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Dat_Id 
      Caption         =   "Dat_Id"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Cmd_Source 
      Caption         =   "..."
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
      Left            =   6960
      TabIndex        =   17
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton Cmd_Famille 
      Caption         =   "..."
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
      Left            =   6960
      TabIndex        =   16
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox Txt_Source 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   15
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox Txt_IdSource 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   14
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Txt_Famille 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   12
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox txt_IdFam 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   11
      Top             =   2880
      Width           =   615
   End
   Begin VB.ComboBox Cmb_Bloque 
      Height          =   315
      ItemData        =   "DocumentWord_fiche.frx":030A
      Left            =   2880
      List            =   "DocumentWord_fiche.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   8
      ToolTipText     =   "La source de données est-elle bloquée ?"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Txt_nom 
      Height          =   285
      Left            =   2880
      MaxLength       =   30
      TabIndex        =   7
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox Txt_Ident 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.CommandButton Cmd_ModifierWord 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   360
         Picture         =   "DocumentWord_fiche.frx":0322
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Modifier le document"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmd_Annuler 
         Height          =   375
         Left            =   7080
         Picture         =   "DocumentWord_fiche.frx":062C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   7800
         Picture         =   "DocumentWord_fiche.frx":072E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_FicheSource 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4740
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12383
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "Source"
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Famille"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Fixe"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Nom"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Identifiant"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "DocumentWord_fiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmd_Annuler_Click()
  Unload Me
End Sub

Private Sub Cmd_Famille_Click()

  Load FamilleDoc_Liste
  With FamilleDoc_Liste
    .cmd_Creer.Visible = False
    .cmd_Modifier.Visible = False
    .cmd_Supprimer.Visible = False
    .cmd_Valider.Visible = True
    .Show vbModal
  End With

End Sub


Private Sub Cmd_ModifierWord_Click()
  MousePointer = vbHourglass
  DocumentWord_liste.ModifierDoc
  MousePointer = vbDefault
End Sub

Private Sub Cmd_Source_Click()

  Load SourcesDonnées_liste
    With SourcesDonnées_liste
      .cmd_Creer.Visible = False
      .cmd_Modifier.Visible = False
      .cmd_Supprimer.Visible = False
      .cmd_Valider.Visible = True
      .Show vbModal
    End With
    
End Sub


Private Sub cmd_Valider_Click()

Dim Fichier As String
Dim CheminFichier As String
Dim sql As String
Dim trouve As Boolean

On Error GoTo Rupture_Integrite

  Select Case DocumentWord_liste.stb_ListeDoc.Panels(4).Text
    Case "CREATION"
      MousePointer = vbHourglass
      sql = "insert into wdoc values(" + Iifent(Txt_Ident) + ","
      sql = sql + Iifch(Txt_nom) + ","
      If Cmb_Bloque.Text = "Oui" Then
        sql = sql + "1" + ","
      Else
        sql = sql + "0" + ","
      End If
      sql = sql + Iifent(txt_IdFam) + ","
      sql = sql + Iifent(Txt_IdSource) + ")"
      Db.Execute sql, dbFailOnError
      DocumentWord_liste.Dat_wdoc.Refresh
      trouve = False
      DocumentWord_liste.Dat_wdoc.Recordset.MoveFirst
      While Not trouve And Not DocumentWord_liste.Dat_wdoc.Recordset.EOF
        If DocumentWord_liste.Dat_wdoc.Recordset.Fields(0) = Iifent(Txt_Ident) Then
          trouve = True
        Else
          DocumentWord_liste.Dat_wdoc.Recordset.MoveNext
        End If
      Wend
      
      DocumentWord_liste.CreerDoc
      MousePointer = vbDefault
    
    Case "MODIFICATION"
      sql = "update wdoc set "
      If Cmb_Bloque.Text = "Oui" Then
        sql = sql + "wdoc_bloque=1"
      Else
        sql = sql + "wdoc_bloque=0"
      End If
      sql = sql + ",wfam_id=" + Iifent(txt_IdFam)
      sql = sql + ",wsou_id=" + Iifent(Txt_IdSource)
      sql = sql + " where wdoc_id=" + CStr(Txt_Ident)
      Workspaces(0).Databases(0).Execute sql, dbFailOnError
      DocumentWord_liste.Dat_wdoc.Refresh
      
    Case "SUPPRESSION"
      sql = "delete * from wdoc where wdoc_id=" + CStr(Txt_Ident)
      Workspaces(0).Databases(0).Execute sql, dbFailOnError
      DocumentWord_liste.Dat_wdoc.Refresh
      If Right(CheminWord, 1) = "\" Then
        CheminFichier = CheminWord + Txt_nom + ".doc"
      Else
        CheminFichier = CheminWord + "\" + Txt_nom + ".doc"
      End If
      Kill CheminFichier
    
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
  MousePointer = vbDefault
Exit Sub

End Sub


Private Sub Form_Load()
  
  Dat_Id.DatabaseName = CheminBase
  
End Sub





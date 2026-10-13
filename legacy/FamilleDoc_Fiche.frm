VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FamilleDoc_Fiche 
   Caption         =   "Famille de documents Fiche"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   ControlBox      =   0   'False
   Icon            =   "FamilleDoc_Fiche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   8775
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Cmb_Bloque 
      Height          =   315
      ItemData        =   "FamilleDoc_Fiche.frx":0442
      Left            =   3240
      List            =   "FamilleDoc_Fiche.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   9
      ToolTipText     =   "La source de données est-elle bloquée ?"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Txt_nom 
      Height          =   285
      Left            =   3240
      MaxLength       =   30
      TabIndex        =   8
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox Txt_id 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.Frame fra_Outils 
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.CommandButton cmd_Valider 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   7800
         Picture         =   "FamilleDoc_Fiche.frx":045A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Annuler 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   7080
         Picture         =   "FamilleDoc_Fiche.frx":055C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_FicheSource 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3555
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12409
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Fixe"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Nom "
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Identifiant"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "FamilleDoc_Fiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Annuler_Click()
  Unload Me
End Sub

Private Sub cmd_Valider_Click()

Dim sql As String
On Error GoTo Rupture_Integrite

  Select Case FamilleDoc_Liste.stb_ListeFamilles.Panels(4).Text
    Case "CREATION"
      sql = "insert into wfam values(" + Iifent(Txt_id) + ","
      sql = sql + Iifch(Txt_Nom) + ","
      If Cmb_Bloque.Text = "Oui" Then
        sql = sql + "1" + ")"
      Else
        sql = sql + "0" + ")"
      End If
    
    Case "MODIFICATION"
      sql = "update wfam set wfam_nom=" + Iifch(Txt_Nom) + ","
      If Cmb_Bloque.Text = "Oui" Then
        sql = sql + "wfam_bloque=1"
      Else
        sql = sql + "wfam_bloque=0"
      End If
      sql = sql + " where wfam_id=" + CStr(Txt_id)
    
    Case "SUPPRESSION"
      sql = "delete * from wfam where wfam_id=" + CStr(Txt_id)
    
  End Select
  
  Workspaces(0).Databases(0).Execute sql, dbFailOnError
  FamilleDoc_Liste.Dat_fam.Refresh
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



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_TypeRecoursFiche 
   Caption         =   "Type de recours contencieux"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   ControlBox      =   0   'False
   Icon            =   "Frm_TypeRecoursFiche.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   7230
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   5640
         Picture         =   "Frm_TypeRecoursFiche.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   6360
         Picture         =   "Frm_TypeRecoursFiche.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_CommuneFiche 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3165
      Width           =   7230
      _ExtentX        =   12753
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
            TextSave        =   "10:34"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6077
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type de recours "
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   6975
      Begin VB.TextBox Txt_code 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Txt_lib 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Code"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Libellé"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   1320
         Width           =   495
      End
   End
End
Attribute VB_Name = "Frm_TypeRecoursFiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Quitter_Click()
  Unload Me
End Sub

Private Sub cmd_Valider_Click()
Const pkstrNomProcédure As String = "cmd_Valider_Click"
On Error GoTo Err_cmd_Valider_Click

Dim requete As String

  MousePointer = vbHourglass
  Select Case Frm_TypeRecoursListe.stb_Types.Panels(4).Text
  
      Case "CREATION"
          requete = "INSERT INTO TYRE VALUES ("
          requete = requete + Iifent(Txt_code) + ","
          requete = requete + Iifch(Txt_lib) + ")"
          
      Case "MODIFICATION"
          requete = "UPDATE TYRE"
          requete = requete + " SET TYRE_NOM=" + Iifch(Txt_lib)
          requete = requete + " WHERE TYRE_NUM=" + Iifent(Txt_code)
      
      Case "SUPPRESSION"
          requete = "DELETE * FROM TYRE WHERE TYRE_NUM = " + Iifent(Txt_code)
          
  End Select
  
'   Ouvrir une transaction
'       Note: L'imbrication de la méthode Execute dans une transaction permet, au moment du CommitTrans, d'enregistrer effectivement toutes les modifications sur disque et de lever les verrous (voir l'aide en ligne sur la méthode Execute). La réimplémentation de la collection des recours contentieux, qui fait appel à une connexion ODBCDirect, dispose alors des valeurs mises à jour.
    Workspaces(0).BeginTrans
    Workspaces(0).Databases(0).Execute requete, dbFailOnError
    Workspaces(0).CommitTrans
    
  MousePointer = vbDefault
    
'   Réimplémenter la collection des recours contentieux
    modDAO.ImplémenterCollections gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_2 & CStr(vbInformation), "Recours contentieux"
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub
    
  Unload Me
  Frm_TypeRecoursListe.Dat_tyre.Refresh
    
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
    Workspaces(0).Rollback
'   S'il s'agit de l'ancienne gestion d'erreurs
    If pbooAncienneGestionErreurs Then
        MousePointer = vbDefault
        Exit Sub
'   Sinon
    Else
        If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    End If

End Sub



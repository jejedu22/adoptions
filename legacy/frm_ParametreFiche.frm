VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_ParametreFiche 
   Caption         =   "Fiche paramètre de l'échéancier"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   7965
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Dat_code 
      Caption         =   "Dat_code"
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
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Txt_nbmois 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Txt_lib 
      Height          =   285
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1800
      Width           =   6255
   End
   Begin VB.TextBox Txt_code 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   6960
         Picture         =   "frm_ParametreFiche.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   6240
         Picture         =   "frm_ParametreFiche.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_EcheancierFiche 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3225
      Width           =   7965
      _ExtentX        =   14049
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
            TextSave        =   "10:38"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7382
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Lbl_nbmois 
      Caption         =   "Nombre de mois"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Lbl_lib 
      Caption         =   "Libellé"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lbl_code 
      Caption         =   "Code"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frm_ParametreFiche"
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
        
    Select Case frm_ParametreListe.stb_ParametreListe.Panels(3).Text
        Case "CREATION"
            requete = "INSERT INTO PARA(PARA_CODE,PARA_LIB,PARA_NBMOIS) VALUES ("
            requete = requete + Iifent(Txt_code) + ","
            requete = requete + Iifch(Txt_lib) + ","
            requete = requete + Iifent(Txt_nbmois) + ");"
            
        Case "MODIFICATION"
            requete = "UPDATE PARA"
            requete = requete + " SET PARA_NBMOIS=" + Iifent(Txt_nbmois)
            requete = requete + " WHERE PARA_CODE=" + Iifent(Txt_code)
            
        Case "SUPPRESSION"
            requete = "DELETE * FROM PARA WHERE PARA_CODE = " + Iifent(Txt_code)
            
    End Select
    
'   Ouvrir une transaction
'       Note: L'imbrication de la méthode Execute dans une transaction permet, au moment du CommitTrans, d'enregistrer effectivement toutes les modifications sur disque et de lever les verrous (voir l'aide en ligne sur la méthode Execute). La réimplémentation de la collection des paramètres de l'échéancier, qui fait appel à une connexion ODBCDirect, dispose alors des valeurs mises à jour.
    Workspaces(0).BeginTrans
    Workspaces(0).Databases(0).Execute requete, dbFailOnError
    Workspaces(0).CommitTrans
    
'   Réimplémenter la collection des paramètres de l'échéancier
    modDAO.ImplémenterCollections gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_2 & CStr(vbInformation), "Paramètres de l'échéancier"
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub
    
    Unload Me
    frm_ParametreListe.Dat_para.Refresh
    
    Exit Sub
  
Err_cmd_Valider_Click:
Dim pbooAncienneGestionErreurs As Boolean
Dim pintTraitement As MessagesActions
    Select Case Err
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case 3022:          pbooAncienneGestionErreurs = True: MsgBox "Ce paramètre existe déjà, création impossible", vbOKOnly + vbExclamation
        Case 3200:          pbooAncienneGestionErreurs = True: MsgBox "Ce paramètre est référencé dans une autre table, suppression impossible", vbOKOnly + vbExclamation
        Case 3075:          pbooAncienneGestionErreurs = True: MsgBox "l'un des champs contient une cote, mise à jour impossible", vbOKOnly + vbExclamation
        Case 3315:          pbooAncienneGestionErreurs = True: MsgBox "Un champs obligatoire n'est pas renseigné, mise à jour impossible", vbOKOnly + vbExclamation
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


Private Sub Form_Load()
  
  Dat_code.DatabaseName = CheminBase
  
End Sub



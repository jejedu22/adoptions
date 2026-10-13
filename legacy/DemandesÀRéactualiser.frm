VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDemandesÀRéactualiser 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   6210
   Begin VB.CommandButton cmdAdoptions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Rafraîchir la liste"
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   6200
   End
   Begin MSComctlLib.ListView lvwAdoptions 
      Height          =   1400
      Index           =   0
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Liste des demandes"
      Top             =   0
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   2461
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N° dossier"
         Object.Width           =   1799
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "N° demande"
         Object.Width           =   1879
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nom"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Prénom"
         Object.Width           =   2910
      EndProperty
   End
End
Attribute VB_Name = "frmDemandesÀRéactualiser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum IndexCommandButton 'Index des contrôles CommandButton
    cmdRAFRAÎCHIR = 0
End Enum
Private Enum IndexListView      'Index des contrôles ListView
    lvwDEMANDES = 0
End Enum

Private Const mstrTITRE As String = "Liste des demandes à réactualiser: "


Private Sub cmdAdoptions_Click(Index As Integer)
Const pkstrNomProcédure As String = "cmdAdoptions_Click"
On Error GoTo Err_cmdAdoptions_Click
    
    Select Case Index
'       Rafraîchir la liste des demandes
        Case cmdRAFRAÎCHIR
            RafraîchirListe
    End Select
        
    Exit Sub
    
Err_cmdAdoptions_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


Private Sub Form_Load()
Const pkstrNomProcédure As String = "Form_Load"
Dim pstrFichierBitmap As String, pstrErreur As String
Dim pbooErreurPossible As Boolean, pbooFichierManquant As Boolean
On Error GoTo Err_Form_Load

'   Positionner la feuille
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) - adpDEMANDES_À_RÉACTUALISER_DÉCALAGE_BAS
    
'   Rafraîchir la liste des demandes
    RafraîchirListe
        
'   Alimenter le contrôle CommandButton
    cmdAdoptions(cmdRAFRAÎCHIR).Picture = Menu.picImage(adpTEXTE_IMAGE, adpTEXTE_RAFRAÎCHIR)

'   Définir l'icône en mode réduit
    pstrFichierBitmap = App.Path & "\" & gstrDOSSIER_IMAGES & "\" & gstrICONE_LISTE
    pbooErreurPossible = True
    Me.Icon = LoadPicture(pstrFichierBitmap)
    If pbooFichierManquant Then Err.Raise mesIMAGE_MANQUANTE
    pbooErreurPossible = False
        
    Exit Sub
    
Err_Form_Load:
Dim pintTraitement As MessagesActions
    If pbooErreurPossible And (Err.Number = comERREUR_00053 Or Err.Number = comERREUR_00068 Or Err.Number = comERREUR_00071 Or Err.Number = comERREUR_00076) Then pbooFichierManquant = True: pstrErreur = Format(Err.Number) & " - " & Err.Description: Resume Next
    Select Case Err.Number
        Case vbObjectError:         pintTraitement = mesINTERROMPRE
        Case mesIMAGE_MANQUANTE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, pstrFichierBitmap, pstrErreur)
        Case Else:                  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Afficher la fenêtre s'il existe une ou plusieurs demandes, décharger la feuille sinon

Public Sub GérerAffichageFenêtre(Optional pbooInformerSiPasÀRéactualiser As Boolean = False)
Const pkstrNomProcédure As String = "GérerAffichageFenêtre"
On Error GoTo Err_GérerAffichageFenêtre

'   S'il n'existe pas de demande à réactualiser
    If lvwAdoptions(lvwDEMANDES).ListItems.Count = 0 Then
    
'       Si l'utilisateur doit en être averti
        If pbooInformerSiPasÀRéactualiser Then
            Err.Raise mesMESSAGE_003
        End If
        
'       Décharger la feuille
        Unload Me
    
    Else
    
'       Afficher la fenêtre
        Me.Show
        
    End If
    
    Exit Sub

Err_GérerAffichageFenêtre:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:     pintTraitement = mesINTERROMPRE
        Case mesMESSAGE_003:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
        Case Else:              pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub



'Rafraîchir la liste des demandes

Private Sub RafraîchirListe()
Const pkstrNomProcédure As String = "RafraîchirListe"
Dim prstDemandesÀRéactualiser As Recordset
Dim pitmDemande As ListItem
On Error GoTo Err_RafraîchirListe

    Screen.MousePointer = vbHourglass
    
    lvwAdoptions(lvwDEMANDES).ListItems.Clear

'   Lire la liste des demandes à réactualiser
    Set prstDemandesÀRéactualiser = modDAO.rstSQLExtraction(daoEXTRACTION_009, False, "La recherche des demandes à réactualiser est annulée.")
    If modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT Then Me.Caption = mstrTITRE & "0"
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Screen.MousePointer = vbDefault: Exit Sub
    
    With prstDemandesÀRéactualiser
'       Pour chacune
        While Not .EOF
'           Créer un item sans image
            Set pitmDemande = lvwAdoptions(lvwDEMANDES).ListItems.Add(, , ![DOSS_NUM])
            modCommun.ImplémenterSubItems pitmDemande, ![DEMA_NUM], modCommun.varValeurFormatéeSiNull(![INDI_NOM], vbString), modCommun.varValeurFormatéeSiNull(![INDI_PREN], vbString)
            .MoveNext
        Wend
        Me.Caption = mstrTITRE & CStr(.RecordCount)
    End With
    
    If Me.Visible Then lvwAdoptions(lvwDEMANDES).SetFocus
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

Err_RafraîchirListe:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


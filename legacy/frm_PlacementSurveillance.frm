VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_PlacementSurveillance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Surveillances du placement"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraAdoptions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   13030
      Begin VB.CommandButton cmdAdoptions 
         Height          =   375
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Créer une surveillance"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAdoptions 
         Height          =   375
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Modifier la surveillance"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAdoptions 
         Height          =   375
         Index           =   2
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Supprimer la surveillance"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.TextBox txtAdoptions 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   11095
         TabIndex        =   16
         Text            =   "0 surveillance(s)"
         ToolTipText     =   "Nombre d'enfants demandés"
         Top             =   540
         Width           =   1695
      End
      Begin MSComctlLib.ListView lvwAdoptions 
         Height          =   2450
         Index           =   0
         Left            =   240
         TabIndex        =   20
         ToolTipText     =   "Historique des surveillances"
         Top             =   960
         Width           =   12550
         _ExtentX        =   22146
         _ExtentY        =   4313
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "<Numéro>"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Demandée le"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Envoi TGI le"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Assistante sociale"
            Object.Width           =   4873
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Intervention le"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Rapport le"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Psychologue"
            Object.Width           =   4445
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Intervention le"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Rapport le"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Historique des surveillances"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   2400
      End
   End
   Begin VB.Frame fraAdoptions 
      Height          =   945
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   13030
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   615
         Left            =   12155
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Valider"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   11435
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Annuler"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame fraAdoptions 
      Caption         =   "Identification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   13030
      Begin VB.TextBox txtAdoptions 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4185
         MaxLength       =   50
         TabIndex        =   11
         Top             =   360
         Width           =   8605
      End
      Begin VB.TextBox txtAdoptions 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtAdoptions 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtAdoptions 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   0
         Left            =   4170
         TabIndex        =   13
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58654723
         CurrentDate     =   2958465
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Déposée le"
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   14
         Top             =   765
         Width           =   810
      End
      Begin VB.Label lblAdoptions 
         AutoSize        =   -1  'True
         Caption         =   "Demandeur"
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   12
         Top             =   405
         Width           =   825
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Placement n°"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   10
         Top             =   1110
         Width           =   945
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Demande n°"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   750
         Width           =   885
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dossier n°"
         Height          =   195
         Index           =   0
         Left            =   405
         TabIndex        =   6
         Top             =   390
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stb_Surveillance 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   6825
      Width           =   13260
      _ExtentX        =   23389
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
            TextSave        =   "17/07/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11:09"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15284
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_PlacementSurveillance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mfrm_PlacementAjoutTS As frm_PlacementAjoutTS
Attribute mfrm_PlacementAjoutTS.VB_VarHelpID = -1

Private Enum IndexColumnHeader          'Index des contrôles ColumnHeader
    hdrNUMÉRO = 0
    hdrDEMANDÉE
    hdrTGI
    hdrASSISTANTE_SOCIALE_NOM
    hdrASSISTANTE_SOCIALE_INTERVENTION
    hdrASSISTANTE_SOCIALE_REÇU
    hdrPSYCHOLOGUE_NOM
    hdrPSYCHOLOGUE_INTERVENTION
    hdrPSYCHOLOGUE_REÇU
End Enum
Private Enum IndexCommandButton 'Index des contrôles CommandButton
    cmdCRÉER = 0
    cmdMODIFIER
    cmdSUPPRIMER
End Enum
Private Enum IndexDTPicker              'Index des contrôles DTPicker
    dtpDÉPOSÉ = 0
End Enum
Private Enum IndexFrame                 'Index des contrôles Frame
    fraOPÉRATIONS = 0
    fraIDENTIFICATION
    fraSURVEILLANCES
End Enum
Private Enum IndexListView              'Index des contrôles ListView
    lvwSURVEILLANCES = 0
End Enum
Private Enum IndexTextBox               'Index des contrôles TextBox
    txtDOSSIER = 0
    txtDEMANDEUR
    txtDEMANDE
    txtPLACEMENT
    txtNOMBRE
End Enum
Private Enum Divers                     'Constantes diverses
    NOMBRE_LIGNES_VISIBLES = 10
End Enum

Private Const mstrNOMBRE_SURVEILLANCES As String = " surveillance(s)"

Private mrstSurveillances As Recordset
Private msngLargeurCadre As Single
Private mlngNuméroPlacement As Long

Private Sub cmd_Quitter_Click()
    Unload Me
End Sub

Private Sub cmd_Valider_Click()
       
  Unload Me
    
End Sub

Private Sub cmdAdoptions_Click(Index As Integer)
Const pkstrNomProcédure As String = "cmdAdoptions_Click"
Dim pitmSurveillanceSélectionnée As ListItem
Dim pstrMessages(0) As String
Dim plngTraitement As CommunTraitements
On Error GoTo Err_cmdAdoptions_Click

'   Référencer la surveillance sélectionnée
    Set pitmSurveillanceSélectionnée = lvwAdoptions(lvwSURVEILLANCES).SelectedItem
    
    Select Case Index
    
'       Créer la surveillance
        Case cmdCRÉER
        
'           Ouvrir la fenêtre permettant le traitement de la surveillance
            With mfrm_PlacementAjoutTS
                .Initialiser comCRÉATION, CLng(txtAdoptions(txtPLACEMENT))
                .Show vbModal
            End With
    
'       Modifier ou supprimer la surveillance
        Case cmdMODIFIER, cmdSUPPRIMER
        
            plngTraitement = IIf(Index = cmdMODIFIER, comMODIFICATION, comSUPPRESSION)
            
'           Refuser la mise à jour si aucune surveillance n'est sélectionnée
            If pitmSurveillanceSélectionnée Is Nothing Then
                pstrMessages(0) = IIf(plngTraitement = comMODIFICATION, "Mise à jour", "Suppression"): Err.Raise mesSÉLECTION_MANQUANTE
                Exit Sub
            End If
        
'           Ouvrir la fenêtre permettant le traitement de la surveillance
            With mfrm_PlacementAjoutTS
                .Initialiser plngTraitement, CLng(txtAdoptions(txtPLACEMENT)), pitmSurveillanceSélectionnée.SubItems(hdrDEMANDÉE), CLng(pitmSurveillanceSélectionnée.Text)
                .Show vbModal
            End With
            
    End Select
        
    Exit Sub
    
Err_cmdAdoptions_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:             pintTraitement = mesINTERROMPRE
        Case mesSÉLECTION_MANQUANTE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, "d'une surveillance", pstrMessages(0) & " refusée")
        Case Else:                      pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const pkstrNomProcédure As String = "Form_Unload"
On Error GoTo Err_Form_Unload

    Set mrstSurveillances = Nothing
    Set mfrm_PlacementAjoutTS = Nothing
    
    Exit Sub
    
Err_Form_Unload:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub

Private Sub mfrm_PlacementAjoutTS_Validation(plngTraitement As CommunTraitements, pdtmSurveillance As Date, pintNuméroRapport As Integer)
Const pkstrNomProcédure As String = "mfrm_PlacementAjoutTS_Validation"
Dim pitmSurveillance As ListItem
On Error GoTo Err_mfrm_PlacementAjoutTS_Validation

'   Implémenter l'historique des surveillances
    Implémenter "Historique des surveillances"
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub
    
'   S'il ne s'agit pas d'une suppression
    If Not (plngTraitement = comSUPPRESSION Or CLng(pintNuméroRapport) = daoENREGISTREMENT_SUPPRIMÉ) Then
'       Référencer la surveillance sur lequel a porté l'opération
        Set pitmSurveillance = lvwAdoptions(lvwSURVEILLANCES).ListItems(CStr(pdtmSurveillance) & gstrSÉPARATEUR_1 & CStr(pintNuméroRapport))
        pitmSurveillance.Selected = True
    End If
    
    Exit Sub
    
Err_mfrm_PlacementAjoutTS_Validation:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Implémenter les objets
'   pstrListeObjets:        liste des objets à implémenter au format '<objet#1>[|<objet#2>[...|<objet#n]]'
'   pstrListeParamètres:    liste des paramètres au format '<paramètre#1 de l'objet#1[¤paramètre#2 de l'objet#1[...¤paramètre#n de l'objet#1]]>[|<paramètre#1 de l'objet#2[¤paramètre#2 de l'objet#2[...¤paramètre#n de l'objet#2]]>[...|<paramètre#1 de l'objet#n[¤paramètre#2 de l'objet#n[...¤paramètre#n de l'objet#n]]>]]'

Private Sub Implémenter(pstrListeObjets As String, Optional pstrListeParamètres As String = "")
Const pkstrNomProcédure As String = "Implémenter"
Dim pitmSurveillance As ListItem
Dim pstrListesObjets() As String, pstrListesParamètres() As String, pstrParamètres() As String, pstrAssistanteSocialeNom As String, pstrPsychologueNom As String
Dim pintItération As Integer
Dim pbooVerrouillé As Boolean
On Error GoTo Err_Implémenter
    
    pstrListesObjets() = Split(pstrListeObjets, ","): pstrListesParamètres() = Split(pstrListeParamètres, gstrSÉPARATEUR_1)
    
'   Pour chaque objet dont on souhaite l'implémentation
    For pintItération = 0 To UBound(pstrListesObjets)
    
        Select Case pstrListesObjets(pintItération)
        
'           Gestion de l'activation des frames
            Case "<désactivation des frames>"

                pstrParamètres() = Split(pstrListesParamètres(pintItération), gstrSÉPARATEUR_2)
                pbooVerrouillé = CBool(pstrParamètres(0))

                modCommun.DéfinirVerrouillageContrôle fraAdoptions(fraSURVEILLANCES), pbooVerrouillé

'           Historique des surveillances
            Case "Historique des surveillances"
        
                lvwAdoptions(lvwSURVEILLANCES).ListItems.Clear
                
'               Si les surveillances n'ont pas encore été lues
                If mrstSurveillances Is Nothing Then
'                   Les lire
                    Set mrstSurveillances = modDAO.rstSQLExtraction(daoSURVEILLANCES, True, gstrCONSULTATION_LIMITÉE, dbOpenSnapshot, CLng(txtAdoptions(txtPLACEMENT)))
'               Sinon
                Else
'                   Rafraîchir le jeu d'enregistrements
                    modDAO.ActualiserRecordset mrstSurveillances, True, gstrCONSULTATION_LIMITÉE
                End If
                If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
                    modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
                    Exit Sub
                End If

                With mrstSurveillances
                
'                   Pour chacun
                    While Not .EOF
                    
                        pstrAssistanteSocialeNom = modAdoptions.txtNomIntervenant(![TRSO_NUM_AssistanteSociale], gstrCONSULTATION_LIMITÉE)
                        pstrPsychologueNom = modAdoptions.txtNomIntervenant(![TRSO_NUM_Psychologue], gstrCONSULTATION_LIMITÉE)
                        
'                       Créer un item sans image
                        Set pitmSurveillance = lvwAdoptions(lvwSURVEILLANCES).ListItems.Add(, CStr(![SUPL_DDSUR]) & gstrSÉPARATEUR_1 & CStr(![SUPL_intNumero]), ![SUPL_intNumero])
                        modCommun.ImplémenterSubItems pitmSurveillance, ![SUPL_DDSUR], modCommun.varValeurFormatéeSiNull(![SUPL_DENVTGI], vbDate, ""), pstrAssistanteSocialeNom, modCommun.varValeurFormatéeSiNull(![SUPL_dtmAssistanteSocialeIntervention], vbDate, ""), modCommun.varValeurFormatéeSiNull(![SUPL_dtmAssistanteSocialeRecu], vbDate, ""), pstrPsychologueNom, modCommun.varValeurFormatéeSiNull(![SUPL_dtmPsychologueIntervention], vbDate, ""), modCommun.varValeurFormatéeSiNull(![SUPL_dtmPsychologueRecu], vbDate, "")
                        
                        .MoveNext
                        
                    Wend
                    
                    txtAdoptions(txtNOMBRE) = CStr(.RecordCount) & mstrNOMBRE_SURVEILLANCES
                    
                End With
        
                lvwAdoptions(lvwSURVEILLANCES).ColumnHeaders(hdrASSISTANTE_SOCIALE_NOM + 1).Width = msngLargeurCadre - IIf(lvwAdoptions(lvwSURVEILLANCES).ListItems.Count > NOMBRE_LIGNES_VISIBLES, comLARGEUR_ASCENCEUR, 0)
        
'           Interdiction de valider
            Case "<interdiction de valider>"
        
                pstrParamètres() = Split(pstrListesParamètres(pintItération), gstrSÉPARATEUR_2)
                pbooVerrouillé = CBool(pstrParamètres(0))
                
                modCommun.DéfinirVerrouillageContrôle cmd_Valider, pbooVerrouillé
                
                If pbooVerrouillé Then
'                   Désactiver les frames
                    Implémenter "<désactivation des frames>", "Vrai"
                End If

        End Select
        
    Next pintItération
    
    Exit Sub
    
Err_Implémenter:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Initialiser la fenêtre
'   plngNuméroPlacement:    numéro du placement

Public Sub Initialiser(plngNuméroPlacement As Long)
Const pkstrNomProcédure As String = "Initialiser"
Dim prstPlacement As Recordset
On Error GoTo Err_Initialiser
    
    mlngNuméroPlacement = plngNuméroPlacement

'   Centrer la feuille
    modAdoptions.CentrerFeuille Me

'   Implémenter les images
    cmd_Quitter.Picture = Menu.picImage(adp32x32, adpANNULER)
    cmd_Valider.Picture = Menu.picImage(adp32x32, adpVALIDER)
    cmdAdoptions(cmdCRÉER).Picture = Menu.picImage(adp16x16, adpCRÉER)
    cmdAdoptions(cmdMODIFIER).Picture = Menu.picImage(adp16x16, adpMODIFIER)
    cmdAdoptions(cmdSUPPRIMER).Picture = Menu.picImage(adp16x16, adpSUPPRIMER)

'   Désactiver certains contrôles
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtDOSSIER), True
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtDEMANDEUR), True
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtDEMANDE), True
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtPLACEMENT), True
    modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpDÉPOSÉ), True

    Set mfrm_PlacementAjoutTS = New frm_PlacementAjoutTS
    msngLargeurCadre = lvwAdoptions(lvwSURVEILLANCES).ColumnHeaders(hdrASSISTANTE_SOCIALE_NOM + 1).Width

'   Lire les références du placement
    Set prstPlacement = modDAO.rstSQLExtraction(daoPLACEMENT_RÉFÉRENCES, True, gstrCONSULTATION_LIMITÉE, mlngNuméroPlacement)
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
        If modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT Then Err.Raise mesENREGISTREMENT_SUPPRIMÉ
'       Interdire la validation
        Implémenter "<interdiction de valider>", "Vrai"
        Exit Sub
    End If
    
    With prstPlacement
        txtAdoptions(txtDOSSIER) = ![DOSS_NUM]
        txtAdoptions(txtDEMANDE) = ![DEMA_NUM]
        txtAdoptions(txtPLACEMENT) = CStr(mlngNuméroPlacement)
        dtpAdoptions(dtpDÉPOSÉ) = ![DEMA_DATE]
    End With
    
'   Afficher l'identité du demandeur
    modAdoptions.AfficherIdentitéDemandeur txtAdoptions(txtDOSSIER), txtAdoptions(txtDEMANDEUR), gstrCONSULTATION_LIMITÉE, cmd_Valider
        
'   Implémenter l'historique des surveillances
    Implémenter "Historique des surveillances"
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub
    
    Exit Sub
    
Err_Initialiser:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                 pintTraitement = mesINTERROMPRE
        Case mesENREGISTREMENT_SUPPRIMÉ:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, "Consultation", "Le placement")
        Case Else:                          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEnvoiDossierTGI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envoi du dossier au TGI"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   3690
   Begin VB.Frame fraAdoptions 
      Height          =   945
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmdAdoptions 
         Default         =   -1  'True
         Height          =   615
         Index           =   1
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Valider"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton cmdAdoptions 
         Cancel          =   -1  'True
         Height          =   615
         Index           =   0
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Annuler"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar stbAdoptions 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2550
      Width           =   3690
      _ExtentX        =   6509
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
            TextSave        =   "04/08/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "13:58"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpAdoptions 
      Height          =   285
      Index           =   0
      Left            =   1100
      TabIndex        =   0
      ToolTipText     =   "Date de l'envoi du dossier au Tribunal de Grande Instance"
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd/MM/yyy"
      Format          =   58916867
      CurrentDate     =   2958465
   End
   Begin VB.Label lblAdoptions 
      Alignment       =   2  'Center
      Caption         =   "Date de l'envoi"
      Height          =   195
      Index           =   0
      Left            =   1100
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "frmEnvoiDossierTGI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum IndexCommandButton 'Index des contrôles CommandButton
    cmdANNULER = 0
    cmdVALIDER
End Enum
Private Enum IndexDTPicker      'Index des contrôles DTPicker
    dtpENVOI = 0
End Enum

Private mrstDateEnvoi As Recordset

Private Sub cmdAdoptions_Click(Index As Integer)
Const pkstrNomProcédure As String = "cmdAdoptions_Click"
On Error GoTo Err_cmdAdoptions_Click

    Select Case Index
    
'       Valider
        Case cmdVALIDER
        
'           Forcer si nécessaire l'actualisation de la valeur d'un contrôle DTPicker
            modCommun.GérerActualisationDTPicker Me, cmdAdoptions(cmdVALIDER)

            With mrstDateEnvoi
            
                .Edit

'               Mettre à jour la date d'envoi
                ![PLAC_dtmEnvoiTGI] = dtpAdoptions(dtpENVOI).Value
                
'               Valider la mise à jour
                Screen.MousePointer = vbHourglass
                modDAO.ValiderOpérationSurRecordset mrstDateEnvoi
                Screen.MousePointer = vbDefault
                Select Case modDAO.lngDernièreErreurDAO()
'                   L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                    Case daoENREGISTREMENT_VERROUILLÉ
'                   L'opération a été refusée en raison de la suppression par un autre utilisateur des données originales
                    Case dbRecordDeleted
'                   L'opération a été refusée en raison de la modification par un autre utilisateur des données originales
                    Case dbRecordModified
'                   L'opération a généré un conflit inattendu
                    Case Is <> daoAUCUNE:   Err.Raise mesCONFLIT_ACCÈS
                End Select
                                
            End With
        
    End Select
        
    Unload Me
        
    Exit Sub
    
Err_cmdAdoptions_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:     pintTraitement = mesINTERROMPRE
        Case mesCONFLIT_ACCÈS:  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, modDAO.lngDernièreErreurDAO())
        Case Else:              pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub

Private Sub Form_Unload(Cancel As Integer)
Const pkstrNomProcédure As String = "Form_Unload"
On Error GoTo Err_Form_Unload

    Set mrstDateEnvoi = Nothing
    
    Exit Sub
    
Err_Form_Unload:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        

End Sub


'Initialiser la fenêtre
'   plngNuméroPlacement:    numéro de placement

Public Sub Initialiser(plngNuméroPlacement As Long)
Const pkstrNomProcédure As String = "Initialiser"
On Error GoTo Err_Initialiser
    
'   Centrer la feuille
    modAdoptions.CentrerFeuille Me
        
'   Implémenter les images
    cmdAdoptions(cmdANNULER).Picture = Menu.picImage(adp32x32, adpANNULER)
    cmdAdoptions(cmdVALIDER).Picture = Menu.picImage(adp32x32, adpVALIDER)
        
'   Lire la date d'envoi
    Set mrstDateEnvoi = modDAO.rstSQLExtraction(daoEXTRACTION_010, True, "La saisie est impossible." & gstrSÉPARATEUR_1 & "Le placement n'existe plus.", plngNuméroPlacement)
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
        modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
        Screen.MousePointer = vbDefault: Exit Sub
    End If
        
'   La positionner
    With mrstDateEnvoi
        dtpAdoptions(dtpENVOI).Value = modCommun.varValeurFormatéeSiNull(![PLAC_dtmEnvoiTGI], vbDate, Date)
        If IsNull(![PLAC_dtmEnvoiTGI]) Then dtpAdoptions(dtpENVOI) = ""
    End With
            
    Exit Sub
    
Err_Initialiser:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHistoriqueCommissions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historique des décisions"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6270
   Begin VB.Frame fraAdoptions 
      Height          =   945
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmdAdoptions 
         Cancel          =   -1  'True
         Default         =   -1  'True
         Height          =   615
         Index           =   0
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fermer"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView lvwAdoptions 
      Height          =   1395
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   960
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Commission"
         Object.Width           =   1799
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Avis"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Signature"
         Object.Width           =   1843
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Décision"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Fratrie"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Motif"
         Object.Width           =   3096
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbAdoptions 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2355
      Width           =   6270
      _ExtentX        =   11060
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
            TextSave        =   "24/07/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "14:38"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4833
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmHistoriqueCommissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum IndexCommandButton 'Index des contrôles CommandButton
    cmdFERMER = 0
End Enum
Private Enum IndexListView      'Index des contrôles ListView
    lvwHISTORIQUE = 0
End Enum


Private Sub cmdAdoptions_Click(Index As Integer)
Const pkstrNomProcédure As String = "cmdAdoptions_Click"
On Error GoTo Err_cmdAdoptions_Click
    
    Select Case Index
'       Fermer la fenêtre
        Case cmdFERMER
            Unload Me
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


'Initialiser la fenêtre
'   plngNuméroDemande:  numéro de demande

Public Sub Initialiser(plngNuméroDemande As Long)
Const pkstrNomProcédure As String = "Initialiser"
Dim prstHistoriqueCommissions As Recordset
Dim pitmCommission As ListItem
On Error GoTo Err_Initialiser

    Screen.MousePointer = vbHourglass
    
'   Centrer la feuille
    modAdoptions.CentrerFeuille Me
        
'   Implémenter les images
    cmdAdoptions(cmdFERMER).Picture = Menu.picImage(adp32x32, adpFERMER)

    lvwAdoptions(lvwHISTORIQUE).ListItems.Clear

'   Lire l'historique des décisions
    Set prstHistoriqueCommissions = modDAO.rstSQLExtraction(daoEXTRACTION_014, False, gstrLECTURE_HISTORIQUE_ANNULÉE, plngNuméroDemande)
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Screen.MousePointer = vbDefault: Exit Sub
    
    With prstHistoriqueCommissions
'       Pour chacune
        While Not .EOF
'           Créer un item sans image
            Set pitmCommission = lvwAdoptions(lvwHISTORIQUE).ListItems.Add(, , ![u02_dtmHistoCommissionReunion])
            modCommun.ImplémenterSubItems pitmCommission, gclsDélibérations(CStr(![u02_intHistoCommissionAvis])).strLibellé, modCommun.varValeurFormatéeSiNull(![u02_dtmHistoCommissionSignature], vbDate, ""), gclsDélibérations(CStr(![u02_intHistoCommissionDecision])).strLibellé, IIf(![u02_booHistoCommissionFratrie], "x", ""), modCommun.varValeurFormatéeSiNull(![u02_txtHistoCommissionMotif], vbString)
            .MoveNext
        Wend
    End With
    
    Screen.MousePointer = vbDefault
            
    Exit Sub
    
Err_Initialiser:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

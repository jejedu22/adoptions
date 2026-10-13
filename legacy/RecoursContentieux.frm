VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRecoursContentieux 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recours contentieux - "
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7065
   Begin VB.Frame fraAdoptions 
      Caption         =   "Décision"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   6855
      Begin VB.TextBox txtAdoptions 
         Height          =   285
         Index           =   2
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   19
         ToolTipText     =   "Motif du refus"
         Top             =   840
         Width           =   5415
      End
      Begin VB.ComboBox cboAdoptions 
         Height          =   315
         Index           =   1
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Nature de la décision"
         Top             =   360
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   15
         ToolTipText     =   "Date de la décision"
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59441155
         CurrentDate     =   2958465
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   21
         ToolTipText     =   "Date de délivrance de l'agrément"
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59441155
         CurrentDate     =   2958465
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Agrément le"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   22
         Top             =   1125
         Width           =   840
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Motif"
         Height          =   195
         Index           =   6
         Left            =   735
         TabIndex        =   18
         Top             =   870
         Width           =   345
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rendue le"
         Height          =   195
         Index           =   4
         Left            =   345
         TabIndex        =   14
         Top             =   405
         Width           =   735
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nature"
         Height          =   195
         Index           =   5
         Left            =   4455
         TabIndex        =   16
         Top             =   405
         Width           =   480
      End
   End
   Begin VB.Frame fraAdoptions 
      Caption         =   "Recours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   6855
      Begin VB.ComboBox cboAdoptions 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Instance auprès de laquelle le recours est déposé"
         Top             =   720
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker dtpAdoptions 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   10
         ToolTipText     =   "Date de dépôt du recours"
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   59441155
         CurrentDate     =   2958465
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Auprès de"
         Height          =   195
         Index           =   3
         Left            =   375
         TabIndex        =   11
         Top             =   765
         Width           =   720
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Déposé le"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   405
         Width           =   720
      End
   End
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
      Height          =   1305
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6855
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
         Width           =   1575
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
         Index           =   1
         Left            =   1200
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Demande n°"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   750
         Width           =   885
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Décision du"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   390
         Width           =   840
      End
   End
   Begin VB.Frame fraAdoptions 
      Height          =   945
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton cmdAdoptions 
         Default         =   -1  'True
         Height          =   615
         Index           =   1
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Valider"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton cmdAdoptions 
         Cancel          =   -1  'True
         Height          =   615
         Index           =   0
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   1
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
      TabIndex        =   20
      Top             =   5475
      Width           =   7065
      _ExtentX        =   12462
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
            TextSave        =   "08/10/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16:28"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6235
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRecoursContentieux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'   Note: La date de délivrance de l'agrément par le Président du Conseil Général n'a de sens que si:
'         - le recours se rapporte à la première commission à prendre une décision, celle-ci ayant décidé de refuser l'agrément
'         - le recours n'est pas déposé auprès du Conseil d'État (CE) après que la Cour Administrative d'Appel (CAA) a pris une décision favorable (dans ce cas, et si la décision du CE confirme celle de la CAA, la date d'agrément saisie à l'occasion du recours déposé auprès de la CAA reste la date d'agrément initial)
'         - la décision est favorable

Public Event Validation(plngTraitement As CommunTraitements, plngNuméroRecours As Long)

Private Enum IndexComboBox      'Index des contrôles ComboBox
    cboAUPRÈS = 0
    cboNATURE
End Enum
Private Enum IndexCommandButton 'Index des contrôles CommandButton
    cmdANNULER = 0
    cmdVALIDER
End Enum
Private Enum IndexDTPicker      'Index des contrôles DTPicker
    dtpDÉPOSÉ = 0
    dtpRENDUE
    dtpAGRÉMENT
End Enum
Private Enum IndexFrame         'Index des contrôles Frame
    fraACTIONS = 0
    fraREFUS_RETRAIT
    fraRECOURS
    fraDÉCISION
End Enum
Private Enum IndexLabel         'Index des contrôles Label
    lblDÉCISION = 0
    lblDEMANDE
    lblDÉPOSÉ
    lblAUPRÈS
    lblRENDUE
    lblNATURE
    lblMOTIF
    lblAGRÉMENT
End Enum
Private Enum IndexTextBox       'Index des contrôles TextBox
    txtDÉCISION = 0
    txtDEMANDE
    txtMOTIF
End Enum

Private mvarRecours As Variant
Private mrstRecours As Recordset
Private mdtmDécision As Date, mdtmSignature As Date
Private mlngDécisionCommission As AdoptionsDélibérations, mlngTraitement As CommunTraitements
Private mlngNuméroDemande As Long, mlngNuméroRecours As Long
Private mbooPremièreCommissionÀPrendreDécision As Boolean, mbooAccordCourAdministrativeAppel As Boolean

Private Sub cboAdoptions_Click(Index As Integer)
Const pkstrNomProcédure As String = "cboAdoptions_Click"
Dim pbooRecoursSurRefusPremièreCommissionÀPrendreDécision As Boolean, pbooRecoursAuprèsCEAprèsDécisionFavorableCAA As Boolean, pbooDécisionFavorable As Boolean, pbooAfficherAgrément As Boolean, pbooRefus As Boolean
On Error GoTo Err_cboAdoptions_Click

    Select Case Index
    
        Case cboNATURE

            pbooRefus = (cboAdoptions(cboNATURE).ItemData(cboAdoptions(cboNATURE).ListIndex) = adpREFUS)
            
'           Activer/Désactiver le motif selon la nature de la décision
            modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtMOTIF), Not pbooRefus
            
'           Afficher/Masquer le motif selon la nature de la décision
            lblAdoptions(lblMOTIF).Visible = pbooRefus: txtAdoptions(txtMOTIF).Visible = pbooRefus
            If Not pbooRefus Then txtAdoptions(txtMOTIF) = ""
            
'           Déterminer si la date de délivrance de l'agrément doit être affichée ou masquée (voir la note en entête de la feuille frmRecoursContentieux)
            pbooRecoursSurRefusPremièreCommissionÀPrendreDécision = mbooPremièreCommissionÀPrendreDécision And (mlngDécisionCommission = adpREFUS)
            pbooRecoursAuprèsCEAprèsDécisionFavorableCAA = ((cboAdoptions(cboAUPRÈS).ItemData(cboAdoptions(cboAUPRÈS).ListIndex) = gclsRecoursContentieux.lngConseilÉtat) And mbooAccordCourAdministrativeAppel)
            pbooDécisionFavorable = (cboAdoptions(cboNATURE).ItemData(cboAdoptions(cboNATURE).ListIndex) = adpACCORD)
            pbooAfficherAgrément = pbooRecoursSurRefusPremièreCommissionÀPrendreDécision And Not pbooRecoursAuprèsCEAprèsDécisionFavorableCAA And pbooDécisionFavorable
            
'           Afficher/Masquer l'agrément
            lblAdoptions(lblAGRÉMENT).Visible = pbooAfficherAgrément: dtpAdoptions(dtpAGRÉMENT).Visible = pbooAfficherAgrément
            If Not pbooAfficherAgrément Then dtpAdoptions(dtpAGRÉMENT) = Date: dtpAdoptions(dtpAGRÉMENT) = ""

    End Select

    Exit Sub
    
Err_cboAdoptions_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
  
End Sub

Private Sub cmdAdoptions_Click(Index As Integer)
Const pkstrNomProcédure As String = "cmdAdoptions_Click"
Dim pvarDécision As Variant
Dim pstrMessages(1) As String, pstrValidationAnnulée As String, pstrDécisionAvantMiseÀJour As String
Dim plngDélibération As AdoptionsDélibérations
On Error GoTo Err_cmdAdoptions_Click

    Select Case Index

'       Valider
        Case cmdVALIDER
        
'           Forcer si nécessaire l'actualisation de la valeur d'un contrôle DTPicker
            modCommun.GérerActualisationDTPicker Me, cmdAdoptions(cmdVALIDER)
    
'           Rafraîchir le jeu des recours contentieux
            modDAO.ActualiserRecordset mrstRecours, True, gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_2 & CStr(vbInformation)
            If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub

'           S'il s'agit d'une suppression
            If mlngTraitement = comSUPPRESSION Then
            
'               Si la suppression n'est pas cohérente
                If modAdoptions.booSaisieIncohérenteDécision(mlngNuméroDemande, gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_2 & CStr(vbInformation), mlngTraitement, mdtmDécision, mlngNuméroRecours, dtpAdoptions(dtpRENDUE), cboAdoptions(cboNATURE), dtpAdoptions(dtpAGRÉMENT), pdtpRecoursDépôt:=dtpAdoptions(dtpDÉPOSÉ), pdtmRefusRetraitSignature:=mdtmSignature, pstrRefusRetrait:=LCase(fraAdoptions(fraREFUS_RETRAIT).Caption), pcmdÀDésactiver:=cmdAdoptions(cmdVALIDER)) Then
'                   Publier l'invitation à actualiser
                    RaiseEvent Validation(comCONSULTATION, mlngNuméroRecours)
                    Exit Sub
                End If

'               Demander confirmation de la suppression du recours
                Select Case modMessages.lngRéponseUtilisateur(mesSUPPRIMER, "le recours contentieux")
'                   L'utilisateur confirme la suppression
                    Case vbYes
'                   L'utilisateur renonce à la suppression
                    Case vbNo
                        cmdAdoptions(cmdANNULER).SetFocus
                        Exit Sub
                End Select
                
'               Si le recours à supprimer existe bien
                If booPositionnementRecoursEffectué() Then

                    With mrstRecours
                
'                       Ouvrir une transaction
                        modDAO.GérerTransaction daoOUVERTURE
    
'                       Valider l'opération
                        modDAO.ValiderOpérationSurRecordset mrstRecours, pbooSupprimer:=True, pstrListeMessages:=gstrMESSAGE_SUR_VERROUILLAGE_EN_SUPPRESSION
                        Select Case modDAO.lngDernièreErreurDAO()
'                           L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                            Case daoENREGISTREMENT_VERROUILLÉ
                                modDAO.GérerTransaction daoANNULATION: Exit Sub
'                           L'opération a généré un conflit inattendu
                            Case Is <> daoAUCUNE
                                modDAO.GérerTransaction daoANNULATION
                                Err.Raise mesCONFLIT_ACCÈS
                        End Select
                        
                    End With

                End If
                
'           Sinon
            Else
            
'               Vérifier la cohérence de la décision
                If modAdoptions.booSaisieIncohérenteDécision(mlngNuméroDemande, gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_2 & CStr(vbInformation), mlngTraitement, mdtmDécision, mlngNuméroRecours, dtpAdoptions(dtpRENDUE), cboAdoptions(cboNATURE), dtpAdoptions(dtpAGRÉMENT), pdtpRecoursDépôt:=dtpAdoptions(dtpDÉPOSÉ), pdtmRefusRetraitSignature:=mdtmSignature, pstrRefusRetrait:=LCase(fraAdoptions(fraREFUS_RETRAIT).Caption), pcmdÀDésactiver:=cmdAdoptions(cmdVALIDER)) Then
'                   Publier l'invitation à actualiser
                    RaiseEvent Validation(comSUPPRESSION, mlngNuméroRecours)
                    Exit Sub
                End If

                With mrstRecours
                
                    Select Case mlngTraitement
                    
'                       Création du recours
                        Case comCRÉATION
                        
'                           Ouvrir une transaction
                            modDAO.GérerTransaction daoOUVERTURE
    
                            .AddNew

                            ![DEMA_NUM] = mlngNuméroDemande
                            ![u02_dtmHistoCommissionReunion] = mdtmDécision
                            ![TYRE_NUM] = mlngNuméroRecours
                            
'                       Mise à jour du recours
                        Case comMODIFICATION
                        
'                           Si le recours à mettre à jour n'existe plus, sortir
                            If Not booPositionnementRecoursEffectué() Then Exit Sub
                            
'                           Déterminer la nature initiale de la décision avant sa mise à jour éventuelle
                            pstrDécisionAvantMiseÀJour = modCommun.varValeurFormatéeSiNull(![RECO_DECI], vbString)
                    
'                           Ouvrir une transaction
                            modDAO.GérerTransaction daoOUVERTURE
    
                            .MoveFirst
                            
'                           Pour chaque recours
                            Do While Not .EOF
                            
'                               S'il s'agit du recours à mettre à jour, sortir de la boucle
'                                   Note: La sortie de la boucle implique que le code qui suit l'instruction ci-dessous ne s'exécutera que pour des recours déposés postérieurement au recours mis à jour
                                If ![TYRE_NUM] = mlngNuméroRecours Then Exit Do
                                
'                               Si le recours mis à jour est déposé auprès de la Cour Administrative d'Appel
                                If cboAdoptions(cboAUPRÈS).ItemData(cboAdoptions(cboAUPRÈS).ListIndex) = gclsRecoursContentieux.lngCourAdministrativeAppel Then
                                
'                                   Si la mise à jour conduit à faire évoluer une décision de refus vers une décision favorable
                                    If pstrDécisionAvantMiseÀJour = gstrREFUS And cboAdoptions(cboNATURE).ItemData(cboAdoptions(cboNATURE).ListIndex) = adpACCORD Then
                                    
                                        .Edit
                                        
'                                       Supprimer la date de délivrance d'agrément éventuellement saisie à l'occasion d'une décision favorable du Conseil d'État (voir la note en entête de la feuille frmRecoursContentieux)
                                        ![RECO_dtmAgrementDelivrance] = Null
                                        
                                        modDAO.ValiderOpérationSurRecordset mrstRecours, pstrListeMessages:=gstrMESSAGE_SUR_VERROUILLAGE_EN_MODIFICATION & gstrSÉPARATEUR_1 & gstrSÉPARATEUR_1 & gstrMESSAGE_SUR_MODIFICATION_SANS_CHOIX, pvarRéponseÀInviteÉcrasement:=vbNo
                                        Select Case modDAO.lngDernièreErreurDAO()
'                                           L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                                            Case daoENREGISTREMENT_VERROUILLÉ
                                                modDAO.GérerTransaction daoANNULATION: Exit Sub
'                                           L'opération a été refusée en raison de la modification par un autre utilisateur des données originales
                                            Case dbRecordModified
                                                modDAO.GérerTransaction daoANNULATION
'                                               Publier l'invitation à actualiser
                                                RaiseEvent Validation(mlngTraitement, daoENREGISTREMENT_SUPPRIMÉ)
                                                modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
                                                Exit Sub
'                                           L'opération a été refusée en raison de la suppression par un autre utilisateur des données originales
                                            Case dbRecordDeleted
'                                           L'opération a généré un conflit inattendu
                                            Case Is <> daoAUCUNE
                                                modDAO.GérerTransaction daoANNULATION
                                                Err.Raise mesCONFLIT_ACCÈS
                                        End Select

'                                       Publier l'invitation à actualiser
                                        RaiseEvent Validation(mlngTraitement, daoENREGISTREMENT_SUPPRIMÉ)
                                        
                                    End If
                                    
                                End If

                                .MoveNext

                            Loop
        
                            .Bookmark = mvarRecours
                            
                            .Edit
                    
                    End Select
                
                    plngDélibération = cboAdoptions(cboNATURE).ItemData(cboAdoptions(cboNATURE).ListIndex): pvarDécision = IIf(plngDélibération = AdoptionsDélibérations.adpINDÉTERMINÉ, Null, IIf(plngDélibération = adpREFUS, gstrREFUS, gstrACCORD))
                    
                    ![RECO_DDREC] = dtpAdoptions(dtpDÉPOSÉ)
                    ![RECO_DPREDEC] = DateAdd("m", modDAO.intParamètreÉchéancier(daoRÉCEPTION_DOSSIER__DÉCISION), ![RECO_DDREC])
                    ![RECO_DDECI] = dtpAdoptions(dtpRENDUE)
                    ![RECO_DECI] = pvarDécision
                    ![RECO_REFUS] = modCommun.varChaîneFormatée(txtAdoptions(txtMOTIF), vbString, Null)
                    ![RECO_dtmAgrementDelivrance] = dtpAdoptions(dtpAGRÉMENT)

'                   Valider l'opération
                    pstrValidationAnnulée = IIf(mlngTraitement = comCRÉATION, gstrMESSAGE_SUR_VERROUILLAGE_EN_CRÉATION, gstrMESSAGE_SUR_VERROUILLAGE_EN_MODIFICATION)
                    modDAO.ValiderOpérationSurRecordset mrstRecours, pstrListeMessages:=pstrValidationAnnulée & gstrSÉPARATEUR_1 & gstrMESSAGE_SUR_SUPPRESSION & gstrSÉPARATEUR_1 & pstrValidationAnnulée & gstrSÉPARATEUR_1 & "Création refusée: Le recours a déjà été déposé."
                    Select Case modDAO.lngDernièreErreurDAO()
#If Not TablesEnLocalOuAttachées Then
'                       La création a été refusée du fait de l'ouverture, par un autre utilisateur, de la base ou de l'objet en mode lecture seule
                        Case daoERREUR_01809
                            modDAO.GérerTransaction daoANNULATION
                            modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
                            Exit Sub
#End If
'                       L'opération a été abandonnée par l'utilisateur après un conflit de verrouillage
                        Case daoENREGISTREMENT_VERROUILLÉ
                            modDAO.GérerTransaction daoANNULATION: Exit Sub
'                       L'opération a été refusée pour éviter la création d'un doublon
                        Case daoDOUBLON
                            modDAO.GérerTransaction daoANNULATION
'                           Publier l'invitation à actualiser
                            RaiseEvent Validation(mlngTraitement, daoENREGISTREMENT_SUPPRIMÉ)
                            modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
                            Exit Sub
'                       L'opération a été refusée en raison de la modification par un autre utilisateur des données originales
                        Case dbRecordModified
'                       L'opération a été refusée en raison de la suppression par un autre utilisateur des données originales
                        Case dbRecordDeleted
                            modDAO.GérerTransaction daoANNULATION
'                           Publier l'invitation à actualiser
                            RaiseEvent Validation(mlngTraitement, daoENREGISTREMENT_SUPPRIMÉ)
                            modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
                            Exit Sub
'                       L'opération a généré un conflit inattendu
                        Case Is <> daoAUCUNE
                            modDAO.GérerTransaction daoANNULATION
                            Err.Raise mesCONFLIT_ACCÈS
                    End Select
                    
                End With
                
            End If
            
'           Actualiser les données relatives à l'agrément
            modAdoptions.ActualiserDonnéesAgrément mlngNuméroDemande, gstrVALIDATION_ANNULÉE
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub

'           Valider la transaction
            modDAO.GérerTransaction daoVALIDATION

'           Publier la validation
            RaiseEvent Validation(mlngTraitement, mlngNuméroRecours)
    
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

Private Sub dtpAdoptions_Change(Index As Integer)
Const pkstrNomProcédure As String = "dtpAdoptions_Change"
On Error GoTo Err_dtpAdoptions_Change

    Select Case Index
    
'       Déposé le
        Case dtpDÉPOSÉ
        
'           Activer/Désactiver les contrôles se rapportant au recours contentieux
            Implémenter "<activation du recours contentieux>"
            
    End Select
    
    Exit Sub
    
Err_dtpAdoptions_Change:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const pkstrNomProcédure As String = "Form_Unload"
On Error GoTo Err_Form_Unload

    Set mrstRecours = Nothing
    
    Exit Sub
    
Err_Form_Unload:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        

End Sub


'Retourner un booléen précisant si un signet a pu être posé sur le recours
'   pbooImplémenter:    booléen précisant si les contrôles doivent être implémentés à partir du recours

Private Function booPositionnementRecoursEffectué(Optional pbooImplémenter As Boolean = False) As Boolean
Const pkstrNomProcédure As String = "booPositionnementRecoursEffectué"
Dim plngDécision As AdoptionsDélibérations
On Error GoTo Err_booPositionnementRecoursEffectué

    With mrstRecours
    
        If Not .RecordCount = 0 Then .MoveFirst
        
'       Pour chaque recours
        Do While Not .EOF
        
'           S'il s'agit du recours à mettre à jour ou à supprimer
            If ![TYRE_NUM] = mlngNuméroRecours Then
            
                mvarRecours = .Bookmark
                booPositionnementRecoursEffectué = True
                
'               Si les contrôles doivent être implémentés
                If pbooImplémenter Then
                    If Not IsNull(![RECO_DDREC]) Then dtpAdoptions(dtpDÉPOSÉ) = ![RECO_DDREC]
                    plngDécision = IIf(IsNull(![RECO_DECI]), AdoptionsDélibérations.adpINDÉTERMINÉ, IIf(![RECO_DECI] = gstrREFUS, adpREFUS, adpACCORD))
                    If Not IsNull(![RECO_DDECI]) Then dtpAdoptions(dtpRENDUE) = ![RECO_DDECI]
                    cboAdoptions(cboNATURE).ListIndex = modCommun.varRecherche(cboAdoptions(cboNATURE), False, plngDécision)
                    txtAdoptions(txtMOTIF) = modCommun.varValeurFormatéeSiNull(![RECO_REFUS], vbString)
                    If Not IsNull(![RECO_dtmAgrementDelivrance]) Then dtpAdoptions(dtpAGRÉMENT) = ![RECO_dtmAgrementDelivrance]
                End If
                
                Exit Do
                
            End If
            
            .MoveNext
            
        Loop
        
    End With

'   Si le recours à mettre à jour ou à supprimer n'existe plus
    If Not booPositionnementRecoursEffectué Then
    
'       Publier l'invitation à actualiser
        RaiseEvent Validation(mlngTraitement, daoENREGISTREMENT_SUPPRIMÉ)
        
        Err.Raise mesENREGISTREMENT_SUPPRIMÉ
        modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
            
    End If

    Exit Function
    
Err_booPositionnementRecoursEffectué:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                 pintTraitement = mesINTERROMPRE
        Case mesENREGISTREMENT_SUPPRIMÉ:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, IIf(mlngTraitement = comMODIFICATION, "Mise à jour", "Suppression"), "Le recours")
        Case Else:                          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Function



'Implémenter les objets
'   pstrListeObjets:        liste des objets à implémenter au format '<objet#1>[|<objet#2>[...|<objet#n]]'
'   pstrListeParamètres:    liste des paramètres au format '<paramètre#1 de l'objet#1[¤paramètre#2 de l'objet#1[...¤paramètre#n de l'objet#1]]>[|<paramètre#1 de l'objet#2[¤paramètre#2 de l'objet#2[...¤paramètre#n de l'objet#2]]>[...|<paramètre#1 de l'objet#n[¤paramètre#2 de l'objet#n[...¤paramètre#n de l'objet#n]]>]]'

Private Sub Implémenter(pstrListeObjets As String, Optional pstrListeParamètres As String = "")
Const pkstrNomProcédure As String = "Implémenter"
Dim pstrListesObjets() As String, pstrListesParamètres() As String, pstrParamètres() As String
Dim pintItération As Integer
Dim pbooVerrouillé As Boolean, pbooRefus As Boolean
On Error GoTo Err_Implémenter
    
    pstrListesObjets() = Split(pstrListeObjets, ","): pstrListesParamètres() = Split(pstrListeParamètres, gstrSÉPARATEUR_1)
    
'   Pour chaque objet dont on souhaite l'implémentation
    For pintItération = 0 To UBound(pstrListesObjets)
    
        Select Case pstrListesObjets(pintItération)
        
'           Gestion de l'activation des contrôles se rapportant au recours contentieux
            Case "<activation du recours contentieux>"
        
                If Not cboAdoptions(cboNATURE).ListIndex = comINDÉFINI Then pbooRefus = (cboAdoptions(cboNATURE).ItemData(cboAdoptions(cboNATURE).ListIndex) = adpREFUS)
                
                pbooVerrouillé = IsNull(dtpAdoptions(dtpDÉPOSÉ))
                modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpRENDUE), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle cboAdoptions(cboNATURE), pbooVerrouillé
                modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtMOTIF), pbooVerrouillé Or Not pbooRefus
                
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
'   plngTraitement:     type de traitement
'   plngNuméroDemande:  numéro de demande
'   pdtmDécision:       date de la décision du refus/retrait
'   plngNuméroRecours:  numéro de recours

Public Sub Initialiser(plngTraitement As CommunTraitements, plngNuméroDemande As Long, pdtmDécision As Date, plngNuméroRecours As Long)
Const pkstrNomProcédure As String = "Initialiser"
Dim prstDécision As Recordset, prstPremièreCommissionÀPrendreDécision As Recordset
Dim pstrMessages(2) As String, pstrDécision As String
On Error GoTo Err_Initialiser
    
    mlngTraitement = plngTraitement: mlngNuméroDemande = plngNuméroDemande: mdtmDécision = pdtmDécision: mlngNuméroRecours = plngNuméroRecours
    mvarRecours = Null

'   Centrer la feuille
    modAdoptions.CentrerFeuille Me
        
'   Repositionner les contrôles
    lblAdoptions(lblAGRÉMENT).Top = lblAdoptions(lblMOTIF).Top
    dtpAdoptions(dtpAGRÉMENT).Top = txtAdoptions(txtMOTIF).Top

'   Implémenter les images
    cmdAdoptions(cmdANNULER).Picture = Menu.picImage(adp32x32, adpANNULER)
    cmdAdoptions(cmdVALIDER).Picture = Menu.picImage(adp32x32, adpVALIDER)
        
'   Désactiver certains contrôles
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtDÉCISION), True
    modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtDEMANDE), True
    modCommun.DéfinirVerrouillageContrôle cboAdoptions(cboAUPRÈS), True
    
'   Implémenter le titre
    Me.Caption = Me.Caption & Choose(mlngTraitement, "Création", "", "Modification", "Suppression")
    
'   Vérifier le paramétrage des recours contentieux
    With gclsRecoursContentieux
        If .booParamétrageIncomplet Then
            pstrMessages(0) = .strTribunalAdministratif: pstrMessages(1) = .strCourAdministrativeAppel: pstrMessages(2) = .strConseilÉtat
            Err.Raise mesPARAMÉTRAGE_CONTENTIEUX_INCOMPLET
            modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
            Exit Sub
        End If
    End With

'   Positionner les références du refus ou du retrait
    txtAdoptions(txtDÉCISION) = CStr(mdtmDécision)
    txtAdoptions(txtDEMANDE) = CStr(mlngNuméroDemande)
    
'   Positionner les dates de dépôt, de décision et d'agrément à la date du jour
    modCommun.PositionnerDateJour dtpAdoptions(dtpDÉPOSÉ), dtpAdoptions(dtpRENDUE), dtpAdoptions(dtpAGRÉMENT)
    dtpAdoptions(dtpDÉPOSÉ) = "": dtpAdoptions(dtpRENDUE) = "": dtpAdoptions(dtpAGRÉMENT) = ""
    
'   Lire la décision de la commission
    Set prstDécision = modDAO.rstSQLExtraction(daoEXTRACTION_023, True, gstrCONSULTATION_LIMITÉE & gstrSÉPARATEUR_1 & "La commission n'existe plus.", mlngNuméroDemande, mdtmDécision)
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
        modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
        Exit Sub
    End If
    
    With prstDécision
    
        mlngDécisionCommission = ![u02_intHistoCommissionDecision]
        pstrDécision = gclsDélibérations(CStr(mlngDécisionCommission)).strLibellé
        fraAdoptions(fraREFUS_RETRAIT).Caption = pstrDécision
    
'       Si la date de signature n'est pas saisie
        If IsNull(![u02_dtmHistoCommissionSignature]) Then
            Err.Raise mesDATE_SIGNATURE_ABSENTE
            modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
'       Sinon
        Else
            mdtmSignature = ![u02_dtmHistoCommissionSignature]
        End If
    
    End With
    
'   Implémenter le type de recours
    modAdoptions.ImplémenterListe cboAdoptions(cboAUPRÈS), "Recours contentieux", "", Nothing
    cboAdoptions(cboAUPRÈS).ListIndex = modCommun.varRecherche(cboAdoptions(cboAUPRÈS), False, mlngNuméroRecours)
'   Implémenter la décision
    modAdoptions.ImplémenterListe cboAdoptions(cboNATURE), "Délibérations", "", Nothing, False
    
'   Lire la date de réunion de la première commission à prendre une décision
    Set prstPremièreCommissionÀPrendreDécision = modDAO.rstSQLExtraction(daoEXTRACTION_031, True, gstrCONSULTATION_LIMITÉE, mlngNuméroDemande)
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then
        modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
        Exit Sub
    End If
    
'   Vérifier s'il s'agit de la commission qui fait l'objet du recours
    With prstPremièreCommissionÀPrendreDécision
        If IsNull(![DatePremiereCommission]) Then
            Err.Raise mesENREGISTREMENT_SUPPRIMÉ
            modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
            Exit Sub
        End If
        mbooPremièreCommissionÀPrendreDécision = (mdtmDécision = ![DatePremiereCommission])
    End With
    
'   Vérifier l'existence de recours contentieux
    Set mrstRecours = modDAO.rstSQLExtraction(daoRECOURS_CONTENTIEUX, True, gstrCONSULTATION_LIMITÉE, dbOpenDynaset, mlngNuméroDemande, mdtmDécision)
    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
        modCommun.DéfinirVerrouillageContrôle cmdAdoptions(cmdVALIDER), True
        Exit Sub
    End If
'   S'il en existe
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT Then
        With mrstRecours
'           Pour chacun
            While Not .EOF
'               S'il est déposé auprès de la Cour Administrative d'Appel
                If ![TYRE_NUM] = gclsRecoursContentieux.lngCourAdministrativeAppel Then
                    mbooAccordCourAdministrativeAppel = (modCommun.varValeurFormatéeSiNull(![RECO_DECI], vbString) = gstrACCORD)
                End If
                .MoveNext
            Wend
        End With
    End If
    
'   S'il s'agit d'une création
    If mlngTraitement = comCRÉATION Then
        cboAdoptions(cboNATURE).ListIndex = modCommun.varRecherche(cboAdoptions(cboNATURE), False, AdoptionsDélibérations.adpINDÉTERMINÉ)
'   Sinon
    Else
'       Si le recours à mettre à jour ou à supprimer n'existe plus, sortir
        If Not booPositionnementRecoursEffectué(True) Then Exit Sub
    End If
    
'   Activer/Désactiver les contrôles se rapportant au recours contentieux
    dtpAdoptions_Change dtpDÉPOSÉ
    
'   S'il s'agit d'une suppression
    If mlngTraitement = comSUPPRESSION Then
'       Désactiver la saisie
        modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpDÉPOSÉ), True, True
        modCommun.DéfinirVerrouillageContrôle cboAdoptions(cboAUPRÈS), True, True
        modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpRENDUE), True, True
        modCommun.DéfinirVerrouillageContrôle cboAdoptions(cboNATURE), True, True
        modCommun.DéfinirVerrouillageContrôle txtAdoptions(txtMOTIF), True, True
        modCommun.DéfinirVerrouillageContrôle dtpAdoptions(dtpAGRÉMENT), True, True
    End If
            
    Exit Sub
    
Err_Initialiser:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                         pintTraitement = mesINTERROMPRE
        Case mesENREGISTREMENT_SUPPRIMÉ:            pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, "Consultation", "La commission")
        Case mesDATE_SIGNATURE_ABSENTE:             pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, LCase(pstrDécision))
        Case mesPARAMÉTRAGE_CONTENTIEUX_INCOMPLET:  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, pstrMessages(0), pstrMessages(1), pstrMessages(2))
        Case Else:                                  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

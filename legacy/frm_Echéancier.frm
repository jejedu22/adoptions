VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm_Echéancier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Édition des retards dans l'échéancier"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   12135
   Begin VB.Frame fraAdoptions 
      Caption         =   "Intervalles par rapport à"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7905
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   11895
      Begin VB.Frame fraAdoptions 
         Caption         =   "Autres"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   6
         Left            =   8040
         TabIndex        =   16
         Top             =   480
         Width           =   3480
         Begin VB.CommandButton cmdAdoptions 
            Caption         =   "Deux surveillances de placement avant jugement"
            Height          =   855
            Index           =   10
            Left            =   240
            TabIndex        =   17
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   3015
         End
      End
      Begin VB.Frame fraAdoptions 
         Caption         =   "Demande de renouvellement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   5
         Left            =   4200
         TabIndex        =   13
         Top             =   4800
         Width           =   3480
         Begin VB.CommandButton cmdAdoptions 
            Caption         =   "Réponse du psychologue"
            Height          =   855
            Index           =   9
            Left            =   240
            TabIndex        =   15
            Top             =   1440
            UseMaskColor    =   -1  'True
            Width           =   3015
         End
         Begin VB.CommandButton cmdAdoptions 
            Caption         =   "Réponse du travailleur social"
            Height          =   855
            Index           =   8
            Left            =   240
            TabIndex        =   14
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   3015
         End
      End
      Begin VB.Frame fraAdoptions 
         Caption         =   "Recours administratif"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Index           =   4
         Left            =   4200
         TabIndex        =   9
         Top             =   480
         Width           =   3480
         Begin VB.CommandButton cmdAdoptions 
            Caption         =   "Décision"
            Height          =   855
            Index           =   5
            Left            =   240
            TabIndex        =   10
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   3015
         End
         Begin VB.CommandButton cmdAdoptions 
            Caption         =   "Réponse du psychologue"
            Height          =   855
            Index           =   7
            Left            =   240
            TabIndex        =   12
            Top             =   2400
            UseMaskColor    =   -1  'True
            Width           =   3015
         End
         Begin VB.CommandButton cmdAdoptions 
            Caption         =   "Réponse du travailleur social"
            Height          =   855
            Index           =   6
            Left            =   240
            TabIndex        =   11
            Top             =   1440
            UseMaskColor    =   -1  'True
            Width           =   3015
         End
      End
      Begin VB.Frame fraAdoptions 
         Caption         =   "Réception du dossier"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Index           =   3
         Left            =   360
         TabIndex        =   5
         Top             =   3480
         Width           =   3480
         Begin VB.CommandButton cmdAdoptions 
            Caption         =   "Décision"
            Height          =   855
            Index           =   4
            Left            =   240
            TabIndex        =   8
            Top             =   2760
            UseMaskColor    =   -1  'True
            Width           =   3015
         End
         Begin VB.CommandButton cmdAdoptions 
            Caption         =   "Réponse du psychologue"
            Height          =   855
            Index           =   3
            Left            =   240
            TabIndex        =   7
            Top             =   1440
            UseMaskColor    =   -1  'True
            Width           =   3015
         End
         Begin VB.CommandButton cmdAdoptions 
            Caption         =   "Réponse du travailleur social"
            Height          =   855
            Index           =   2
            Left            =   240
            TabIndex        =   6
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   3015
         End
      End
      Begin VB.Frame fraAdoptions 
         Caption         =   "Demande"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   3480
         Begin VB.CommandButton cmdAdoptions 
            Caption         =   "Réception du dossier"
            Height          =   855
            Index           =   1
            Left            =   240
            TabIndex        =   4
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   3015
         End
      End
   End
   Begin VB.Frame fraAdoptions 
      Height          =   945
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton cmdAdoptions 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   615
         Index           =   0
         Left            =   10920
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fermer"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin Crystal.CrystalReport cryÉtat 
         Left            =   0
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
   End
   Begin MSComctlLib.StatusBar stb_ListeAssurMa 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   9210
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "05/08/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "08:19"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15187
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_Echéancier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum IndexCommandButton                             'Index des contrôles CommandButton
    cmdFERMER = 0
    cmdDEMANDE__RÉCEPTION_DOSSIER
    cmdRÉCEPTION_DOSSIER__RÉPONSE_TRAVAILLEUR_SOCIAL
    cmdRÉCEPTION_DOSSIER__RÉPONSE_PSYCHOLOGUE
    cmdRÉCEPTION_DOSSIER__DÉCISION
    cmdRECOURS__DÉCISION
    cmdRECOURS__RÉPONSE_TRAVAILLEUR_SOCIAL
    cmdRECOURS__RÉPONSE_PSYCHOLOGUE
    cmdDEMANDE_RENOUVELLEMENT__RÉPONSE_TRAVAILLEUR_SOCIAL
    cmdDEMANDE_RENOUVELLEMENT__RÉPONSE_PSYCHOLOGUE
    cmdDEUX_SURVEILLANCES_AVANT_JUGEMENT
End Enum


Private Sub cmdAdoptions_Click(Index As Integer)
Const pkstrNomProcédure As String = "cmdAdoptions_Click"
Dim pdtmBaseCalcul As Date
Dim pstrDateBaseCalcul As String, pstrDateFormatée As String, pstrRequête As String, pstrÉtatCrystalReports As String, pstrErreur As String
Dim pintAttributs As Integer
Dim pbooErreurPossible As Boolean, pbooFichierManquant As Boolean
On Error GoTo Err_cmdAdoptions_Click

'   S'il s'agit de fermer la fenêtre
    If Index = cmdFERMER Then
        Unload Me
        Exit Sub
    End If
    
'   Si l'utilisateur décide de prendre la date du jour comme base de calcul pour l'échéancier
    If modMessages.lngRéponseUtilisateur(mesQUESTION_011) = vbYes Then
        pdtmBaseCalcul = Date
'   Sinon
    Else
'       Proposer la saisie d'une date
        pstrDateBaseCalcul = InputBox("Date à prendre comme base de calcul" & vbCrLf & "pour l'échéancier : " & vbCrLf & vbCrLf & "(Format: JJ/MM/AAAA)")
'       Vérifier sa validité
        If pstrDateBaseCalcul = "" Then Exit Sub
        If IsDate(pstrDateBaseCalcul) Then pdtmBaseCalcul = CDate(pstrDateBaseCalcul) Else Err.Raise mesMESSAGE_021: Exit Sub
    End If

    pstrDateFormatée = Year(pdtmBaseCalcul) & ", " & Month(pdtmBaseCalcul) & ", " & Day(pdtmBaseCalcul)

    Select Case Index
      
'       Demande / Réception du dossier
        Case cmdDEMANDE__RÉCEPTION_DOSSIER
            pstrÉtatCrystalReports = "Dossier"
            pstrRequête = "{INDI.INDI_CDFR} = 'DE' And IsNull({DEMA.DEMA_DSORT}) And IsNull({DEMA.DEMA_DSUSP}) And IsNull({DEMA.DEMA_DREDOS}) And ((Not IsNull({DEMA.DEMA_DPRDOS}) And Date({DEMA.DEMA_DPRDOS}) < DateTime(" & pstrDateFormatée & ")) Or (IsNull({DEMA.DEMA_DPRDOS}) And Not IsNull({DEMA.DEMA_DREF}) And Date({DEMA.DEMA_DREF} + (30 *" & modDAO.intParamètreÉchéancier(daoDEMANDE__RÉCEPTION_DOSSIER) & ")) < DateTime(" & pstrDateFormatée & ")))"

'       Réception du dossier / Réponse du travailleur social
        Case cmdRÉCEPTION_DOSSIER__RÉPONSE_TRAVAILLEUR_SOCIAL
            pstrÉtatCrystalReports = "Travailleur_social"
            pstrRequête = "{INDI.INDI_CDFR} = 'DE' And IsNull({DEMA.DEMA_DSORT}) And IsNull({DEMA.DEMA_DSUSP}) And Not IsNull({DEMA.DEMA_DDEMRA}) And IsNull({DEMA.DEMA_DRECRA}) And ((Not IsNull({DEMA.DEMA_DPRRA}) And {DEMA.DEMA_DPRRA} < DateTime(" & pstrDateFormatée & ")) Or (IsNull({DEMA.DEMA_DPRRA}) And Not IsNull({DEMA.DEMA_DREF}) And {DEMA.DEMA_DREF} + (30 * " & modDAO.intParamètreÉchéancier(daoRÉCEPTION_DOSSIER__RÉPONSE_TRAVAILLEUR_SOCIAL) & ") < DateTime(" & pstrDateFormatée & ")))"

'       Réception du dossier / Réponse du psychologue
        Case cmdRÉCEPTION_DOSSIER__RÉPONSE_PSYCHOLOGUE
            pstrÉtatCrystalReports = "Psychologue"
            pstrRequête = "{INDI.INDI_CDFR} = 'DE' And IsNull({DEMA.DEMA_DSORT}) And IsNull({DEMA.DEMA_DSUSP}) And Not IsNull({DEMA.DEMA_DDEMRP}) And IsNull({DEMA.DEMA_DRECRP}) And ((Not IsNull({DEMA.DEMA_DPRRP}) And {DEMA.DEMA_DPRRP} < DateTime(" & pstrDateFormatée & ")) Or (IsNull({DEMA.DEMA_DPRRP}) And Not IsNull({DEMA.DEMA_DREF}) And {DEMA.DEMA_DREF} + (30 * " & modDAO.intParamètreÉchéancier(daoRÉCEPTION_DOSSIER__RÉPONSE_PSYCHOLOGUE) & ") < DateTime(" & pstrDateFormatée & ")))"

'       Réception du dossier / Décision
        Case cmdRÉCEPTION_DOSSIER__DÉCISION
            pstrÉtatCrystalReports = "Decision"
            pstrRequête = "{INDI.INDI_CDFR} = 'DE' And IsNull({DEMA.DEMA_DSORT}) And IsNull({DEMA.DEMA_DSUSP}) And IsNull({u02_HistoCommissions.u02_dtmHistoCommissionSignature}) And ((Not IsNull({u02_HistoCommissions.u02_dtmHistoCommissionPrevu}) And {u02_HistoCommissions.u02_dtmHistoCommissionPrevu} < DateTime(" & pstrDateFormatée & ")) Or (IsNull({u02_HistoCommissions.u02_dtmHistoCommissionPrevu}) And Not IsNull({DEMA.DEMA_DREF}) And ({DEMA.DEMA_DREF} + (30 * " & modDAO.intParamètreÉchéancier(daoRÉCEPTION_DOSSIER__DÉCISION) & ") < DateTime(" & pstrDateFormatée & "))))"

'       Recours administratif / Décision
        Case cmdRECOURS__DÉCISION
            pstrÉtatCrystalReports = "Decision-Recours"
            pstrRequête = "{INDI.INDI_CDFR} = 'DE' And IsNull({READ.READ_DDECI}) And (Not IsNull({READ.READ_DPREDEC}) And {READ.READ_DPREDEC} < DateTime(" & pstrDateFormatée & "))"

'       Recours administratif / Réponse du travailleur social
        Case cmdRECOURS__RÉPONSE_TRAVAILLEUR_SOCIAL
            pstrÉtatCrystalReports = "Travailleur_Social-recours"
            pstrRequête = "{INDI.INDI_CDFR} = 'DE' And IsNull({DEMA.DEMA_DSORT}) And IsNull({DEMA.DEMA_DSUSP}) And Not IsNull({READ.READ_DDEMAS}) And IsNull({READ.READ_RECAS}) And (Not IsNull({READ.READ_DPRRECAS}) And {READ.READ_DPRRECAS} < DateTime(" & pstrDateFormatée & "))"

'       Recours administratif / Réponse du psychologue
        Case cmdRECOURS__RÉPONSE_PSYCHOLOGUE
            pstrÉtatCrystalReports = "Psychologue-recours"
            pstrRequête = "{INDI.INDI_CDFR} = 'DE' And IsNull({DEMA.DEMA_DSORT}) And IsNull({DEMA.DEMA_DSUSP}) And Not IsNull({READ.READ_DDEMPSY}) And IsNull({READ.READ_RECPSY}) And (Not IsNull({READ.READ_DPRRECPSY}) And {READ.READ_DPRRECPSY} < DateTime(" & pstrDateFormatée & "))"

'       Demande de renouvellement / Réponse du travailleur social
        Case cmdDEMANDE_RENOUVELLEMENT__RÉPONSE_TRAVAILLEUR_SOCIAL
            pstrÉtatCrystalReports = "Renouvellement_AS"
            pstrRequête = "{INDI.INDI_CDFR} = 'DE' And IsNull({DEMA.DEMA_DSORT}) And IsNull({DEMA.DEMA_DSUSP}) And Not IsNull({RENO.RENO_DDEMMAAS}) And IsNull({RENO.RENO_DRECMAS}) And Not IsNull({RENO.RENO_DPRERECAS}) And {RENO.RENO_DPRERECAS} <  DateTime(" & pstrDateFormatée & ")"

'       Demande de renouvellement / Réponse du psychologue
        Case cmdDEMANDE_RENOUVELLEMENT__RÉPONSE_PSYCHOLOGUE
            pstrÉtatCrystalReports = "Renouvellement_Psy"
            pstrRequête = "{INDI.INDI_CDFR} = 'DE' And IsNull({DEMA.DEMA_DSORT}) And IsNull({DEMA.DEMA_DSUSP}) And Not IsNull({RENO.RENO_DDEMMPS}) And IsNull({RENO.RENO_DRECMPS}) And Not IsNull({RENO.RENO_DPRERECPSY}) And {RENO.RENO_DPRERECPSY} < DateTime(" & pstrDateFormatée & ")"

'       Deux surveillances de placement avant jugement
        Case cmdDEUX_SURVEILLANCES_AVANT_JUGEMENT
            pstrÉtatCrystalReports = "Surveillance"
            pstrRequête = "{INDI.INDI_CDFR} = 'DE' And IsNull({DEMA.DEMA_DSORT}) And IsNull({PLAC.JUAC_DJUGTPA}) And ((IsNull({SUPL.SUPL_dtmAssistanteSocialeRecu}) And Not IsNull({SUPL.SUPL_dtmAssistanteSocialePrevu}) And {SUPL.SUPL_dtmAssistanteSocialePrevu} < DateTime(" & pstrDateFormatée & ")) Or (IsNull({SUPL.SUPL_dtmPsychologueRecu}) And Not IsNull({SUPL.SUPL_dtmPsychologuePrevu}) And {SUPL.SUPL_dtmPsychologuePrevu} < DateTime(" & pstrDateFormatée & ")))"

    End Select

'   Vérifier la disponibilité du fichier Crystal Reports
    pstrÉtatCrystalReports = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Crystal\" & pstrÉtatCrystalReports & ".rpt"
    pbooErreurPossible = True
    pintAttributs = GetAttr(pstrÉtatCrystalReports)
    If pbooFichierManquant Then Err.Raise mesMESSAGE_022
    pbooErreurPossible = False
                
    With cryÉtat
    
        .DataFiles(0) = modDAO.strNomBaseAvecChemin()
        .Destination = crptToWindow
        .ReportFileName = pstrÉtatCrystalReports
        
        .ReplaceSelectionFormula (pstrRequête)
        .Action = 1
        
    End With

    Exit Sub
    
Err_cmdAdoptions_Click:
Dim pintTraitement As MessagesActions
    If pbooErreurPossible And (Err.Number = comERREUR_00053 Or Err.Number = comERREUR_00068 Or Err.Number = comERREUR_00071 Or Err.Number = comERREUR_00076) Then pbooFichierManquant = True: pstrErreur = Format(Err.Number) & " - " & Err.Description: Resume Next
    Select Case Err.Number
        Case vbObjectError:     pintTraitement = mesINTERROMPRE
        Case mesMESSAGE_021:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
        Case mesMESSAGE_022:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, "Le fichier nécessaire à la construction de l'état", "", pstrÉtatCrystalReports, pstrErreur)
        Case Else:              pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub


'Initialiser la fenêtre

Public Sub Initialiser()
Const pkstrNomProcédure As String = "Initialiser"
On Error GoTo Err_Initialiser

    Screen.MousePointer = vbHourglass
    
'   Centrer la feuille
    modAdoptions.CentrerFeuille Me
        
'   Implémenter les images
    cmdAdoptions(cmdFERMER).Picture = Menu.picImage(adp32x32, adpFERMER)
    
'   Implémenter les libellés des boutons
    cmdAdoptions(cmdDEUX_SURVEILLANCES_AVANT_JUGEMENT).Caption = "Deux surveillances de placement" & vbCrLf & "avant jugement"

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


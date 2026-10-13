VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_DemandeAdoptionEnfant 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4455
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraAdoptions 
      Caption         =   "Souhaits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   5295
      Begin VB.ComboBox Cmb_Sexe 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Sexe : (M)asculin / (F)éminin / (I)ndifférent"
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox Cmb_Handicap 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Handicap : (O)ui / (N)on"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Txt_trancheAge 
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   14
         ToolTipText     =   "Tranche d'âge de l'enfant demandé (ex: 01-04)"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lbl_sexe 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sexe"
         Height          =   195
         Left            =   1440
         TabIndex        =   9
         Top             =   405
         Width           =   360
      End
      Begin VB.Label Lbl_handicap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Handicap"
         Height          =   195
         Left            =   1110
         TabIndex        =   11
         Top             =   765
         Width           =   690
      End
      Begin VB.Label Lbl_TrancheAges 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tranche d'âges"
         Height          =   195
         Left            =   690
         TabIndex        =   13
         Top             =   1080
         Width           =   1110
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   945
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Annuler"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   615
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Valider"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar stb_DemandeEnfant 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   4200
      Width           =   5520
      _ExtentX        =   9737
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
            TextSave        =   "05/08/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15:41"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3510
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Fra_enfant 
      Caption         =   "Numéros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5295
      Begin VB.TextBox Txt_NumDemande 
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
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Txt_code 
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
         Left            =   1920
         TabIndex        =   7
         ToolTipText     =   "Code du régime d'assurance maladie"
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Enfant"
         Height          =   195
         Index           =   1
         Left            =   1335
         TabIndex        =   6
         Top             =   765
         Width           =   465
      End
      Begin VB.Label lblAdoptions 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Demande"
         Height          =   195
         Index           =   0
         Left            =   1110
         TabIndex        =   4
         Top             =   405
         Width           =   690
      End
      Begin VB.Label Lbl__separateur 
         Caption         =   "/"
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   360
         Width           =   15
      End
   End
End
Attribute VB_Name = "frm_DemandeAdoptionEnfant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Validation(plngTraitement As CommunTraitements, pintNuméroEnfant As Integer)

Private Enum IndexFrame 'Index des contrôles Frame
    fraENFANT = 0
End Enum
Private Enum IndexLabel             'Index des contrôles Label
    lblDEMANDE = 0
    lblENFANT
End Enum

Private mlngTraitement As CommunTraitements
Private mbooEnfantDemandé As Boolean

Private Sub cmd_Quitter_Click()
Const pkstrNomProcédure As String = "cmd_Quitter_Click"
On Error GoTo Err_cmd_Quitter_Click

    Unload Me
    
    Exit Sub
    
Err_cmd_Quitter_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_Valider_Click()
Const pkstrNomProcédure As String = "cmd_Valider_Click"
Dim psrtTable As String
Dim pintNuméroEnfant As Integer
On Error GoTo Err_cmd_Valider_Click

    psrtTable = IIf(mbooEnfantDemandé, "FADE", "COAC")
  
'   Ouvrir une transaction
    modDAO.GérerTransaction daoOUVERTURE
    
    Select Case mlngTraitement
    
'       Création de l'enfant (demandé ou accordé)
        Case comCRÉATION
    
'           Lire le numéro du dernier enfant créé
            pintNuméroEnfant = modDAO.varValeurMaximale("ENDE", "ENDE_NUM", gstrVALIDATION_ANNULÉE)
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
            
'           Créer l'enfant
            Do
                pintNuméroEnfant = pintNuméroEnfant + 1
                modDAO.SQLAjouter daoADD_ENFANT, True, gstrVALIDATION_ANNULÉE & gstrSÉPARATEUR_1 & gstrDOUBLON_SANS_MESSAGE, pintNuméroEnfant, Choose(Cmb_Sexe.ItemData(Cmb_Sexe.ListIndex), "I", "M", "F"), Choose(Cmb_Handicap.ItemData(Cmb_Handicap.ListIndex) + 2, "O", "N"), modCommun.varChaîneFormatée(Txt_trancheAge, vbString, Null)
                If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoDOUBLON) Then modDAO.GérerTransaction daoANNULATION: Exit Sub
            Loop While modDAO.lngDernièreErreurDAO() = daoDOUBLON
    
'           Créer l'enfant demandé ou accordé en commission
            modDAO.SQLAjouter daoADD_ENFANT_DEMANDÉ_OU_ACCORDÉ, True, gstrVALIDATION_ANNULÉE, psrtTable, CLng(Txt_NumDemande), pintNuméroEnfant
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
            
'       Mise à jour de l'enfant (demandé ou accordé)
        Case comMODIFICATION
    
'           Le mettre à jour
            pintNuméroEnfant = CInt(Txt_code)
            modDAO.SQLMettreÀJour daoMISE_À_JOUR_003, True, gstrVALIDATION_ANNULÉE, pintNuméroEnfant, Choose(Cmb_Sexe.ItemData(Cmb_Sexe.ListIndex), "I", "M", "F"), Choose(Cmb_Handicap.ItemData(Cmb_Handicap.ListIndex) + 2, "O", "N"), modCommun.varChaîneFormatée(Txt_trancheAge, vbString, Null)
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
            
'           S'il n'existe plus
            If modDAO.lngNombreEnregistrementsAffectés() = 0 Then
            
                modDAO.GérerTransaction daoANNULATION
                
'               Publier l'invitation à actualiser
                RaiseEvent Validation(mlngTraitement, daoENREGISTREMENT_SUPPRIMÉ)
                
                Err.Raise mesENREGISTREMENT_SUPPRIMÉ
                modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
                Exit Sub
                
            End If
    
'       Mise à jour de l'enfant (demandé ou accordé)
        Case comSUPPRESSION
    
'           Le supprimer
            pintNuméroEnfant = CInt(Txt_code)
            modDAO.SQLSupprimer daoDEL_ORIGINES_OU_ENFANT_OU_DEMANDÉ_OU_ACCORDÉ, True, gstrVALIDATION_ANNULÉE, "ORSO", pintNuméroEnfant
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
            modDAO.SQLSupprimer daoDEL_ORIGINES_OU_ENFANT_OU_DEMANDÉ_OU_ACCORDÉ, True, gstrVALIDATION_ANNULÉE, psrtTable, pintNuméroEnfant
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
            modDAO.SQLSupprimer daoDEL_ORIGINES_OU_ENFANT_OU_DEMANDÉ_OU_ACCORDÉ, True, gstrVALIDATION_ANNULÉE, "ENDE", pintNuméroEnfant
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then modDAO.GérerTransaction daoANNULATION: Exit Sub
    
    End Select
    
'   Valider la transaction
    modDAO.GérerTransaction daoVALIDATION
    
'   Publier la validation
    RaiseEvent Validation(mlngTraitement, pintNuméroEnfant)
    
    Unload Me
    
    Exit Sub
  
Err_cmd_Valider_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                 pintTraitement = mesINTERROMPRE
        Case mesENREGISTREMENT_SUPPRIMÉ:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, "Mise à jour", "L'enfant")
        Case Else:                          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub


'Initialiser la fenêtre
'   plngTraitement:     type de traitement
'   pbooDemande:        booléen précisant si le souhait est exprimé lors de la demande (True) ou lors de la commission (False)
'   plngNuméroDemande:  numéro de demande
'   pvarNuméroEnfant:   numéro d'enfant [obligatoire s'il ne s'agit pas d'une création]

Public Sub Initialiser(plngTraitement As CommunTraitements, pbooDemande As Boolean, plngNuméroDemande As Long, Optional pvarNuméroEnfant As Variant)
Const pkstrNomProcédure As String = "Initialiser"
Dim prstEnfant As Recordset
On Error GoTo Err_Initialiser
    
    mlngTraitement = plngTraitement: mbooEnfantDemandé = pbooDemande
    
'   Centrer la feuille
    modAdoptions.CentrerFeuille Me
        
'   Implémenter les images
    cmd_Quitter.Picture = Menu.picImage(adp32x32, adpANNULER)
    cmd_Valider.Picture = Menu.picImage(adp32x32, adpVALIDER)
    
'   Désactiver certains contrôles
    modCommun.DéfinirVerrouillageContrôle Txt_NumDemande, True
    modCommun.DéfinirVerrouillageContrôle Txt_code, True
    
'   Implémenter le titre
    Me.Caption = "Souhaits lors de la " & IIf(pbooDemande, "demande", "commission") & " - " & Choose(mlngTraitement, "Création", "", "Modification", "Suppression")
    
'   Implémenter le sexe
    modAdoptions.ImplémenterListe Cmb_Sexe, "Sexes", "", Nothing, True
'   Implémenter l'handicap
    modCommun.ImplémenterListeOuiNon Cmb_Handicap
    
'   Positionner les numéros de demande et d'enfant
    Txt_NumDemande = CStr(plngNuméroDemande)
    If Not IsMissing(pvarNuméroEnfant) Then Txt_code = CStr(pvarNuméroEnfant)
    
'   S'il s'agit d'une création
    If mlngTraitement = comCRÉATION Then
    
        Txt_code.Visible = False: lblAdoptions(lblENFANT).Visible = False
        Cmb_Sexe.ListIndex = modCommun.varRecherche(Cmb_Sexe, False, AdoptionsSexes.adpINDÉTERMINÉ)
        Cmb_Handicap.ListIndex = modCommun.varRecherche(Cmb_Handicap, False, False)
        
'   Sinon
    Else
    
'       Lire l'enfant
        Set prstEnfant = modDAO.rstSQLExtraction(daoEXTRACTION_022, True, gstrCONSULTATION_LIMITÉE, CInt(pvarNuméroEnfant))
        If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
            modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
            Exit Sub
        End If
        
'       Si l'enfant à mettre à jour ou à supprimer n'existe plus
        If modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT Then
        
'           Publier l'invitation à actualiser
            RaiseEvent Validation(mlngTraitement, daoENREGISTREMENT_SUPPRIMÉ)
            
            Err.Raise mesENREGISTREMENT_SUPPRIMÉ
            modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
            Exit Sub
            
        End If
        
        With prstEnfant
            Cmb_Sexe.ListIndex = modCommun.varRecherche(Cmb_Sexe, False, IIf(![ENDE_SEXE] = "I", AdoptionsSexes.adpINDÉTERMINÉ, IIf(![ENDE_SEXE] = "M", adpMASCULIN, adpFÉMININ)))
            Cmb_Handicap.ListIndex = modCommun.varRecherche(Cmb_Handicap, False, CLng(![ENDE_HAND] = "O"))
            Txt_trancheAge = modCommun.varValeurFormatéeSiNull(![ENDE_TRAGE], vbString)
        End With
        
    End If
    
'   S'il s'agit d'une suppression
    If mlngTraitement = comSUPPRESSION Then
'       Désactiver la saisie
        modCommun.DéfinirVerrouillageContrôle Txt_trancheAge, True, True
        modCommun.DéfinirVerrouillageContrôle Cmb_Sexe, True, True
        modCommun.DéfinirVerrouillageContrôle Cmb_Handicap, True, True
    End If
            
    Exit Sub
    
Err_Initialiser:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                 pintTraitement = mesINTERROMPRE
        Case mesENREGISTREMENT_SUPPRIMÉ:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, IIf(mlngTraitement = comMODIFICATION, "Mise à jour", "Suppression"), "L'enfant")
        Case Else:                          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


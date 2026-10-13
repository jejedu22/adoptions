VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adoptions"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6450
   ControlBox      =   0   'False
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6450
   Begin PicClip.PictureClip pclAdoptions 
      Index           =   0
      Left            =   120
      Top             =   240
      _ExtentX        =   1508
      _ExtentY        =   873
      _Version        =   393216
   End
   Begin VB.Data Dat_CheminWord 
      Caption         =   "Dat_CheminWord"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Dat_CodeGenerateur 
      Caption         =   "Dat_CodeGenerateur"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin PicClip.PictureClip pclAdoptions 
      Index           =   1
      Left            =   120
      Top             =   720
      _ExtentX        =   1508
      _ExtentY        =   873
      _Version        =   393216
   End
   Begin PicClip.PictureClip pclAdoptions 
      Index           =   2
      Left            =   120
      Top             =   1200
      _ExtentX        =   1508
      _ExtentY        =   873
      _Version        =   393216
   End
   Begin VB.Menu Fichier 
      Caption         =   "&Fichier"
      Begin VB.Menu Quitter 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu Edition 
      Caption         =   "&Edition"
      Enabled         =   0   'False
   End
   Begin VB.Menu Définir 
      Caption         =   "&Définir"
      Begin VB.Menu RegimeASMA 
         Caption         =   "Régimes assurance &maladie"
      End
      Begin VB.Menu sitaufamil 
         Caption         =   "Situations &familiales"
      End
      Begin VB.Menu Professions 
         Caption         =   "&Professions"
      End
      Begin VB.Menu separ0 
         Caption         =   "-"
      End
      Begin VB.Menu Nature 
         Caption         =   "&Natures des demandes/placements"
      End
      Begin VB.Menu Types 
         Caption         =   "T&ypes des demandes"
      End
      Begin VB.Menu TypeRecours 
         Caption         =   "Typ&es des recours contentieux"
      End
      Begin VB.Menu Origines 
         Caption         =   "Origines et&hniques des enfants"
      End
      Begin VB.Menu separ1 
         Caption         =   "-"
      End
      Begin VB.Menu Institutions 
         Caption         =   "&Institutions des travailleurs sociaux"
      End
      Begin VB.Menu Qualité 
         Caption         =   "&Qualités des travailleurs sociaux"
      End
      Begin VB.Menu Permanence 
         Caption         =   "&Lieux de permanences"
      End
   End
   Begin VB.Menu Préparer 
      Caption         =   "&Préparer"
      Begin VB.Menu TS 
         Caption         =   "&Intervenants"
      End
      Begin VB.Menu Oeuvre 
         Caption         =   "&Oeuvres de placements"
      End
      Begin VB.Menu Tribunaux 
         Caption         =   "&Tribunaux"
      End
      Begin VB.Menu separ3 
         Caption         =   "-"
      End
      Begin VB.Menu Communes 
         Caption         =   "Co&mmunes"
      End
      Begin VB.Menu Cantons 
         Caption         =   "Ca&ntons"
      End
      Begin VB.Menu Circonscriptions 
         Caption         =   "Circ&onscriptions"
      End
      Begin VB.Menu separ2 
         Caption         =   "-"
      End
      Begin VB.Menu SOURCESDONNEES 
         Caption         =   "&Sources de données"
      End
      Begin VB.Menu FAMILLES 
         Caption         =   "&Familles de documents"
      End
      Begin VB.Menu DOCUMENTS 
         Caption         =   "&Documents"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu PARAMETRE 
         Caption         =   "Paramètres de l'échéancier"
      End
   End
   Begin VB.Menu Gérer 
      Caption         =   "&Gérer"
      Begin VB.Menu Dossiers 
         Caption         =   "&Gestion des dossiers/demandes"
      End
      Begin VB.Menu separation1 
         Caption         =   "-"
      End
      Begin VB.Menu GestDemandes 
         Caption         =   "Gestion des &demandes"
         Begin VB.Menu Instruction 
            Caption         =   "&Instruction d'une demande"
         End
         Begin VB.Menu mnuSuspendreReprendre 
            Caption         =   "Sus&pendre / Reprendre l'instruction"
         End
         Begin VB.Menu COMMISSION 
            Caption         =   "Décision de &la commission"
         End
         Begin VB.Menu RECOURS 
            Caption         =   "Re&cours"
         End
         Begin VB.Menu CONFIRMATION 
            Caption         =   "C&onfirmation de la demande"
         End
         Begin VB.Menu ECHEANCIER 
            Caption         =   "Redéfinir la date de calcul de l'éc&héancier d'une demande"
         End
      End
      Begin VB.Menu separation2 
         Caption         =   "-"
      End
      Begin VB.Menu GestPlacements 
         Caption         =   "Gestion des &placements"
         Begin VB.Menu ListePlacements 
            Caption         =   "&Liste des placements"
         End
         Begin VB.Menu Placement 
            Caption         =   "&Enregister/afficher un placement"
         End
      End
   End
   Begin VB.Menu Exploiter 
      Caption         =   "E&xploiter"
      Begin VB.Menu ETAT 
         Caption         =   "Édition des retards dans l'é&chéancier"
      End
      Begin VB.Menu MAILINGS 
         Caption         =   "&Mailings"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu CALCULAGE 
         Caption         =   "&Recalcul de l'âge moyen"
      End
      Begin VB.Menu mnuSéparateur 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRechercheDemandesÀRéactualiser 
         Caption         =   "Recherche des &demandes à réactualiser"
      End
   End
   Begin VB.Menu Affichage 
      Caption         =   "Aff&ichage"
      Enabled         =   0   'False
   End
   Begin VB.Menu Options 
      Caption         =   "&Options"
      Enabled         =   0   'False
   End
   Begin VB.Menu Aide 
      Caption         =   "&?"
      Begin VB.Menu A_Propos 
         Caption         =   "A &Propos"
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Matrices           'Constantes permettant la division des bitmaps en matrices uniformes
    LIGNES_TEXTE_IMAGE = 3
    COLONNES_TEXTE_IMAGE = 1
    LIGNES_16x16 = 3
    COLONNES_16x16 = 3
    LIGNES_32x32 = 1
    COLONNES_32x32 = 3
End Enum
Private Enum Divers             'Constantes diverses
    HAUTEUR_FENÊTRE = 700
End Enum

Private Const mstrBITMAP_TEXTE_IMAGE As String = "Texte+Image.bmp", mstrBITMAP_16x16 As String = "16x16.bmp", mstrBITMAP_32x32 As String = "32x32.bmp"

Public Sub CalculAgeMoyen()
'met à jour l'âge moyen d'un couple
Dim RsDossier As Recordset
Dim RsIndi As Recordset
Dim sql As String
Dim NbIndi, NbMois As Integer
Dim An, Mois As Integer

On Error GoTo Rupture_Integrite

  sql = "select * from doss"
  Set RsDossier = Db.OpenRecordset(sql)
  If Not RsDossier.EOF Then
    RsDossier.MoveFirst
    While Not RsDossier.EOF
      NbIndi = 0
      NbMois = 0
      'demandeur
      sql = "select indi_dnais from indi where indi_cdfr=""DE"""
      sql = sql + " and doss_num=" + Iifch(RsDossier("doss_num"))
      Set RsIndi = Db.OpenRecordset(sql)
      If Not RsIndi.EOF Then
        NbIndi = NbIndi + 1
        NbMois = Abs(DateDiff("m", Date, RsIndi("indi_dnais")))
      End If
      RsIndi.Close
      
      'conjoint
      sql = "select indi_dnais from indi where indi_cdfr=""CO"""
      sql = sql + " and doss_num=" + Iifch(RsDossier("doss_num"))
      Set RsIndi = Db.OpenRecordset(sql)
      If Not RsIndi.EOF Then
        NbIndi = NbIndi + 1
        NbMois = NbMois + Abs(DateDiff("m", Date, RsIndi("indi_dnais")))
      End If
      RsIndi.Close
      
      If NbIndi = 0 Then
        MsgBox "Le dossier " + CStr(RsDossier("doss_num")) + " ne contient ni demandeur, ni conjoint !"
      Else
        'calcul An
        An = NbMois / (NbIndi * 12)
        An = Round(An, 0)
        If An * 12 * NbIndi > NbMois Then An = An - 1 'cas arrondi supérieur
        'calcul Mois
        Mois = (NbMois - (An * 12 * NbIndi)) / 2
        
        'modifie l'enregistrement DOSSIER
        sql = "update doss set doss_an=" + CStr(An) + ",doss_mois=" + CStr(Mois)
        sql = sql + " where doss_num =" + Iifch(RsDossier("doss_num"))
        Db.Execute sql, dbFailOnError
      End If
      RsDossier.MoveNext
    Wend
  End If
  RsDossier.Close
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


Private Sub A_Propos_Click()
    frm_Apropos.Show vbModal
End Sub

Private Sub CALCULAGE_Click()
  CalculAgeMoyen
End Sub

Private Sub Cantons_Click()
    frm_Canton.Show vbModal
End Sub

Private Sub Circonscriptions_Click()
    frm_Circonscription.Show vbModal
End Sub



Private Sub COMMISSION_Click()
Const pkstrNomProcédure As String = "COMMISSION_Click"
On Error GoTo Err_COMMISSION_Click

'   Ouvrir la fenêtre permettant le traitement des commissions
    With frm_DecisionCommission
        .Initialiser
        .Show vbModal
    End With

    Exit Sub
    
Err_COMMISSION_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub Communes_Click()
    frm_Commune.Show vbModal
End Sub

Private Sub CONFIRMATION_Click()

  Load Frm_ConfirmationDemandeListe

  Frm_ConfirmationDemandeListe.Txt_NomDemandeur.Enabled = False
  Frm_ConfirmationDemandeListe.Txt_PrenomDemandeur.Enabled = False
  Frm_ConfirmationDemandeListe.DTP_DateDemande.Enabled = False
  Frm_ConfirmationDemandeListe.Show vbModal
  
End Sub

Private Sub DOCUMENTS_Click()
  DocumentWord_liste.Show vbModal
End Sub

Private Sub Dossiers_Click()
    frm_DossierGestion.Show vbModeless
End Sub

Private Sub Echeancier_Click()
Const pkstrNomProcédure As String = "Echeancier_Click"
On Error GoTo Err_Echeancier_Click

'   Ouvrir la fenêtre permettant le recalcul dans l'échéancier
    With frm_RecalculEcheancier
        .Initialiser
        .Show vbModal
    End With

    Exit Sub
    
Err_Echeancier_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub ETAT_Click()
Const pkstrNomProcédure As String = "ETAT_Click"
On Error GoTo Err_ETAT_Click

'   Ouvrir la fenêtre permettant l'édition des retards de l'échéancier
    With frm_Echéancier
        .Initialiser
        .Show
    End With

    Exit Sub
    
Err_ETAT_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub FAMILLES_Click()
  FamilleDoc_Liste.Show vbModal
End Sub

Private Sub Form_Load()
Const pkstrNomProcédure As String = "Form_Load"
On Error GoTo Err_Form_Load

    Dat_CodeGenerateur.DatabaseName = CheminBase
    CentrageHorizontal Me
    
'   Extraire les images des fichiers bitmaps
    ExtraireBitmaps
    
'   Forcer le dimensionnement correct de la fenêtre
'       Note: Sans cette instruction, les menus sont partiellement coupés, avec la valeur Me.Height = 0
    Me.Height = HAUTEUR_FENÊTRE
    Me.Top = 0

    Exit Sub
    
Err_Form_Load:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub Form_Terminate()
Const pkstrNomProcédure As String = "Form_Terminate"
On Error GoTo Err_Form_Terminate

'   Supprimer les classes
    modAdoptions.SupprimerClasses
    
    Exit Sub
    
Err_Form_Terminate:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub Institutions_Click()
    frm_Institution.Show vbModal
End Sub

Private Sub Instruction_Click()
Const pkstrNomProcédure As String = "Instruction_Click"
On Error GoTo Err_Instruction_Click

'   Ouvrir la fenêtre permettant le traitement de l'instruction
    With frm_DemandeAdoptionInstruction
        .Initialiser
        .Show vbModal
    End With

    Exit Sub
    
Err_Instruction_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub ListePlacements_Click()
    frm_PlacementListe.Show vbModeless
End Sub

Private Sub MAILINGS_Click()

  Load DocumentWord_liste
  With DocumentWord_liste
    .cmd_Creer.Visible = False
    .cmd_Modifier.Visible = False
    .cmd_Supprimer.Visible = False
    .Cmd_Mailings.Visible = True
    .Show vbModal
  End With
  
End Sub

Private Sub mnuRechercheDemandesÀRéactualiser_Click()
Const pkstrNomProcédure As String = "mnuRechercheDemandesÀRéactualiser_Click"
On Error GoTo Err_mnuRechercheDemandesÀRéactualiser_Click

'   Lister les demandes à réactualiser
    frmDemandesÀRéactualiser.Hide
    frmDemandesÀRéactualiser.GérerAffichageFenêtre True
    
    Exit Sub
    
Err_mnuRechercheDemandesÀRéactualiser_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub mnuSuspendreReprendre_Click()
Const pkstrNomProcédure As String = "mnuSuspendreReprendre_Click"
On Error GoTo Err_mnuSuspendreReprendre_Click

'   Ouvrir la fenêtre permettant le traitement de la suspension ou de la reprise
    With frm_DemandeAdoptionSuspension
        .Initialiser
        .Show vbModal
    End With

    Exit Sub
    
Err_mnuSuspendreReprendre_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub Nature_Click()
    frm_NatureDemandeEtPlacement.Show vbModal
End Sub

Private Sub Oeuvre_Click()
    frm_Oeuvres.Show vbModal
End Sub

Private Sub Origines_Click()
    frm_OrigineEthnique.Show vbModal
End Sub

Private Sub PARAMETRE_Click()
    frm_ParametreListe.Show vbModeless
End Sub

Private Sub Permanence_Click()
    frm_Permanence.Show vbModal
End Sub

Private Sub Placement_Click()
    
    Load frm_PlacementGestion
    frm_PlacementGestion.stb_Placement.Panels(4).Text = "CREATION"
    Call ListeDossiers(frm_PlacementGestion.Dat_dossiers, frm_PlacementGestion.Cmb_NumDossier)
    Call ListeDemandesPourCeDossier(frm_PlacementGestion.Dat_demandes, frm_PlacementGestion.Cmb_NumDemande, frm_PlacementGestion.Cmb_NumDossier.Text)
    frm_PlacementGestion.Txt_NomDemandeur.Enabled = False
    frm_PlacementGestion.Txt_PrenomDemandeur.Enabled = False
    frm_PlacementGestion.DTP_DateDemande.Enabled = False
    
    frm_PlacementGestion.Txt_NumeroPlacement = NouveauCode("PLAC", "PLAC_NUM")
    frm_PlacementGestion.DTP_DatePlacement = Date
    frm_PlacementGestion.DTP_DateConseilFamille = Date
    frm_PlacementGestion.DTP_DateJugementAccueil = Date
    frm_PlacementGestion.DTP_DateJugementOrigine = Date
    frm_PlacementGestion.DTP_DateConseilFamille = ""
    frm_PlacementGestion.DTP_DateJugementAccueil = ""
    frm_PlacementGestion.DTP_DateJugementOrigine = ""
    frm_PlacementGestion.Show vbModal
End Sub

Private Sub Professions_Click()
    frm_Profession.Show vbModal
End Sub

Private Sub Qualité_Click()
    frm_Qualite.Show vbModal
End Sub

Private Sub Quitter_Click()
    'Fermeture des fenêtres restant chargées en mémoire
    Do While Forms.Count > 0
        Unload Forms(Forms.Count - 1)
    Loop
    'Déconnexion de tous les DATA de la base de données
    Do While Workspaces(0).Databases.Count <> 0
        Do While Workspaces(0).Databases(0).Recordsets.Count <> 0
            Workspaces(0).Databases(0).Recordsets(0).Close
        Loop
        Workspaces(0).Databases(0).Close
    Loop
End Sub

Private Sub RECOURS_Click()
Const pkstrNomProcédure As String = "RECOURS_Click"
On Error GoTo Err_RECOURS_Click

'   Ouvrir la fenêtre permettant le traitement des recours
    With Frm_Recours
        .Initialiser
        .Show vbModal
    End With

    Exit Sub
    
Err_RECOURS_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub RegimeASMA_Click()
    frm_AssurMa.Show vbModal
End Sub


Private Sub sitaufamil_Click()
    frm_SituationFam.Show vbModal
End Sub

Private Sub SOURCESDONNEES_Click()
  SourcesDonnées_liste.Show vbModal
End Sub

Private Sub Tribunaux_Click()
    frm_Tribunaux.Show vbModeless
End Sub

Private Sub TS_Click()
Const pkstrNomProcédure As String = "TS_Click"
On Error GoTo Err_TS_Click

'   Ouvrir la fenêtre permettant le traitement de l'intervenant
    With frm_TS
        .Initialiser
        .Show
    End With

    Exit Sub
    
Err_TS_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub TypeRecours_Click()
  Frm_TypeRecoursListe.Show vbModal
End Sub

Private Sub Types_Click()
    frm_TypeDemande.Show vbModal
End Sub


'Extraire les fichiers bitmaps

Private Sub ExtraireBitmaps()
Const pkstrNomProcédure As String = "ExtraireBitmaps"
Dim pstrDossierImages As String, pstrFichierBitmap As String, pstrErreur As String
Dim pbooErreurPossible As Boolean, pbooFichierManquant As Boolean
On Error GoTo Err_ExtraireBitmaps

    pstrDossierImages = App.Path & "\" & gstrDOSSIER_IMAGES & "\"
    
'   Alimenter les contrôles PictureClip
    pbooErreurPossible = True
    With pclAdoptions(adpTEXTE_IMAGE)
        pstrFichierBitmap = pstrDossierImages & mstrBITMAP_TEXTE_IMAGE
        .Picture = LoadPicture(pstrFichierBitmap)
        If pbooFichierManquant Then Err.Raise mesIMAGE_MANQUANTE
        .Rows = LIGNES_TEXTE_IMAGE: .Cols = COLONNES_TEXTE_IMAGE
    End With
    With pclAdoptions(adp16x16)
        pstrFichierBitmap = pstrDossierImages & mstrBITMAP_16x16
        .Picture = LoadPicture(pstrFichierBitmap)
        If pbooFichierManquant Then Err.Raise mesIMAGE_MANQUANTE
        .Rows = LIGNES_16x16: .Cols = COLONNES_16x16
    End With
    With pclAdoptions(adp32x32)
        pstrFichierBitmap = pstrDossierImages & mstrBITMAP_32x32
        .Picture = LoadPicture(pstrFichierBitmap)
        If pbooFichierManquant Then Err.Raise mesIMAGE_MANQUANTE
        .Rows = LIGNES_32x32: .Cols = COLONNES_32x32
    End With
    pbooErreurPossible = False

    Exit Sub

Err_ExtraireBitmaps:
Dim pintTraitement As MessagesActions
    If pbooErreurPossible And (Err.Number = comERREUR_00053 Or Err.Number = comERREUR_00068 Or Err.Number = comERREUR_00071 Or Err.Number = comERREUR_00076) Then pbooFichierManquant = True: pstrErreur = Format(Err.Number) & " - " & Err.Description: Resume Next
    Select Case Err.Number
        Case vbObjectError:         pintTraitement = mesINTERROMPRE
        Case mesIMAGE_MANQUANTE:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing, pstrFichierBitmap, pstrErreur)
        Case Else:                  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Retourner une image extraite de l'un des contrôles PictureClip
'
'   plngFichierBitmap:  index du contrôle PictureClip
'   plngImage:          index de l'image

Public Function picImage(plngFichierBitmap As AdoptionsIndexPictureClip, plngImage As AdoptionsImages) As Picture
Const pkstrNomProcédure As String = "picImage"
On Error GoTo Err_picImage

    Set picImage = Menu.pclAdoptions(plngFichierBitmap).GraphicCell(plngImage)
    
    Exit Function
    
Err_picImage:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
            
End Function


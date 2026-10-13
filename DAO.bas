Attribute VB_Name = "modDAO"
Option Explicit

Private Const mstrNOM_MODULE As String = "modDAO"

Public Enum DAOErreurs                                      'Erreurs DAO
    daoENREGISTREMENT_VERROUILLÉ = -5                       'Un enregistrement est actuellement verrouillé par une transaction
    daoAUCUN_ENREGISTREMENT                                 'Le jeu ne contient pas d'enregistrement
    daoDOUBLON                                              'L'ajout (ou la mise à jour) a été refusé(e) pour éviter la création d'un doublon
    daoENREGISTREMENT_CONNEXE_MANQUANT                      'L'ajout (ou la mise à jour) a été refusé(e) du fait d'un enregistrement connexe manquant
    daoENREGISTREMENT_CONNEXE_EXISTANT                      'La suppression (ou la mise à jour) a été refusée du fait d'un enregistrement connexe existant
    daoAUCUNE
#If Not TablesEnLocalOuAttachées Then
    daoERREUR_01017 = -1017                                 'S1109: [Microsoft][Pilote ODBC Microsoft Access] L'enregistrement est supprimé.
    daoERREUR_01611 = -1611                                 'S1000: [Microsoft][Pilote ODBC Microsoft Access] Le moteur de base de données Microsoft Jet a arrêté le traitement parce que vous et un autre utilisateur tentez de modifier les mêmes données en même temps.
    daoERREUR_01809 = -1809                                 '42000: [Microsoft][Pilote ODBC Microsoft Access] Mise à jour impossible. La base de données ou l'objet est en lecture seule.
#End If
    daoERREUR_01304 = -1304                                 'S1000: [Microsoft][Pilote ODBC Microsoft Access] Impossible de verrouiller la table '<#1>'; actuellement utilisée par l'utilisateur '<#2>' sur l'ordinateur '<#3>'.
    daoERREUR_03021 = 3021                                  'Pas d'enregistrement courant.
    daoERREUR_03022 = -1605                                 '23000: [Microsoft][Pilote ODBC Microsoft Access] Modifications non effectuées: risque de doublons dans champs index, clé principale ou relation interdisant les doublons. Modifiez les données des champs contenant les doublons, enlevez ou redéfinissez l'index pour permettre les doublons et recommencez.
    daoERREUR_03035 = -3035                                 'S1000: [Microsoft][Pilote ODBC Microsoft Access] Impossible de verrouiller la table '<#1>'; actuellement utilisée par l'utilisateur '<#2>' sur l'ordinateur '<#3>'.
    daoERREUR_03049 = -3049                                 'S1000: [Microsoft][Pilote ODBC Microsoft Access] Impossible de verrouiller la table '<#1>'; actuellement utilisée par l'utilisateur '<#2>' sur l'ordinateur '<#3>'.
    daoERREUR_03146 = 3146                                  'ODBC -- L'appel a échoué.
    daoERREUR_03167 = 3167                                  'L'enregistrement est supprimé.
    daoERREUR_03200 = -1612                                 'S1000: [Microsoft][Pilote ODBC Microsoft Access] Impossible de supprimer ou de modifier l'enregistrement car la table '<#1>' comprend des enregistrements connexes.
    daoERREUR_03201 = -1613                                 '23000: [Microsoft][Pilote ODBC Microsoft Access] Vous ne pouvez pas ajouter ou modifier un enregistrement car l'enregistrement associé est requis dans la table '<#1>'.
    daoERREUR_03260 = -1102                                 'S1000: [Microsoft][Pilote ODBC Microsoft Access] Impossible de mettre à jour; actuellement verrouillé par l'utilisateur '<#1>' sur l'ordinateur '<#2>'.
    daoERREUR_03261 = -1302                                 'S1000: [Microsoft][Pilote ODBC Microsoft Access] La table '<#1>' est verrouillée exclusivement par l'utilisateur '<#2>' sur l'ordinateur '<#3>'.
End Enum
Public Enum DAOExtractions                                  'Extractions de jeux d'enregistrements
    daoINDIVIDU = 1
    daoNOMBRE
    daoMAXIMUM
    daoRECOURS_CONTENTIEUX
    daoPSYCHOLOGUES__OU__ASSISTANTES_SOCIALES
    daoDEMANDE_DATES
    daoRECOURS_ADMINISTRATIF
    daoSURVEILLANCES
    daoDÉCISIONS
    daoENFANTS_DEMANDÉS_OU_ACCORDÉS
    daoPLACEMENT_RÉFÉRENCES
    daoEXTRACTION_001
    daoEXTRACTION_002
    daoEXTRACTION_003
    daoEXTRACTION_004
    daoEXTRACTION_005
    daoEXTRACTION_006
    daoEXTRACTION_007
    daoEXTRACTION_008
    daoEXTRACTION_009
    daoEXTRACTION_010
    daoEXTRACTION_011
    daoEXTRACTION_012
    daoEXTRACTION_013
    daoEXTRACTION_014
    daoEXTRACTION_015
    daoEXTRACTION_016
    daoEXTRACTION_017
    daoEXTRACTION_018
    daoEXTRACTION_019
    daoEXTRACTION_020
    daoEXTRACTION_021
    daoEXTRACTION_022
    daoEXTRACTION_023
    daoEXTRACTION_024
    daoEXTRACTION_025
    daoEXTRACTION_026
    daoEXTRACTION_027
    daoEXTRACTION_028
    daoEXTRACTION_029
    daoEXTRACTION_031
    daoEXTRACTION_032
End Enum
Public Enum DAOAjouts                                       'Ajouts d'enregistrement
    daoADD_RECOURS_ADMINISTRATIF = 1
    daoADD_ENFANT
    daoADD_ENFANT_DEMANDÉ_OU_ACCORDÉ
    daoADD_HISTORIQUE_INTERVENANT
    daoAJOUT_001
    daoAJOUT_002
    daoAJOUT_003
End Enum
Public Enum DAOMisesÀJour                                   'Mises à jour d'enregistrement
    daoMAJ_HISTORIQUE_INTERVENANT = 1
    daoMISE_À_JOUR_001
    daoMISE_À_JOUR_002
    daoMISE_À_JOUR_003
    daoMISE_À_JOUR_004
End Enum
Public Enum DAOSuppressions                                 'Suppressions d'enregistrements
    daoDEL_HISTORIQUE_COMMISSION_OU_RECOURS = 1
    daoDEL_ORIGINES_OU_ENFANT_OU_DEMANDÉ_OU_ACCORDÉ
    daoDEL_HISTORIQUE_INTERVENANT
    daoSUPPRESSION_001
End Enum
Public Enum DAOTransactions                                 'Méthodes de transactions
    daoOUVERTURE = 1
    daoVALIDATION
    daoANNULATION
End Enum
Public Enum DAOParamètresÉchéancier                         'Paramètres de l'échéancier
    daoDEMANDE__RÉCEPTION_DOSSIER = 1
    daoRÉCEPTION_DOSSIER__RÉPONSE_TRAVAILLEUR_SOCIAL
    daoRÉCEPTION_DOSSIER__RÉPONSE_PSYCHOLOGUE
    daoRÉCEPTION_DOSSIER__DÉCISION
    daoDEMANDE_RENOUVELLEMENT__RÉPONSE_TRAVAILLEUR_SOCIAL
    daoDEUX_SURVEILLANCES_AVANT_JUGEMENT
    daoRECOURS_ADMINISTRATIF__RÉPONSE_TRAVAILLEUR_SOCIAL
    daoRECOURS_ADMINISTRATIF__RÉPONSE_PSYCHOLOGUE
End Enum
Public Enum DAODivers                                       'Constantes diverses
    daoENREGISTREMENT_SUPPRIMÉ = -1
End Enum
Public Enum DAOTaillesColonnes                              'Tailles de colonnes
    daoWCOL_NOM = 50
End Enum

Public Const gstrDONNÉES_SUPPRIMÉES As String = "Certaines données ont été supprimées par un utilisateur depuis que vous les utilisez."
Public Const gstrMESSAGE_ARRÊT_SUR_VERROUILLAGE As String = "L'application est arrêtée." & gstrSÉPARATEUR_2 & "16", gstrMESSAGE_SUR_VERROUILLAGE_EN_VALIDATION As String = "La validation est annulée." & gstrSÉPARATEUR_2 & "64", gstrMESSAGE_SUR_VERROUILLAGE_EN_CRÉATION As String = "La création est annulée." & gstrSÉPARATEUR_2 & "64", gstrMESSAGE_SUR_VERROUILLAGE_EN_MODIFICATION As String = "La mise à jour est annulée." & gstrSÉPARATEUR_2 & "64", gstrMESSAGE_SUR_VERROUILLAGE_EN_SUPPRESSION As String = "La suppression est annulée." & gstrSÉPARATEUR_2 & "64"
Public Const gstrMESSAGE_SUR_SUPPRESSION As String = "Mise à jour refusée: " & gstrDONNÉES_SUPPRIMÉES & gstrSÉPARATEUR_2 & "48", gstrMESSAGE_SUR_MODIFICATION_ET_CHOIX_ANNULER As String = "La validation est annulée." & gstrSÉPARATEUR_2 & "64", gstrMESSAGE_SUR_MODIFICATION_SANS_CHOIX As String = "Mise à jour refusée: Certaines données ont été modifiées par un utilisateur depuis que vous les utilisez." & gstrSÉPARATEUR_2 & "48"
Public Const gstrDOUBLON_SANS_MESSAGE As String = "<Pas de message>"

Public gclsProvenances As clsCodesLibellés, gclsRecoursContentieux As clsRecoursContentieux

Private Enum MessagesDAO                                    'Messages DAO
    ERREUR_1304_NOMBRE_APOSTROPHES = 8
    ERREUR_1304_NUMÉRO_APOSTROPHE_AVANT_NOM_UTILISATEUR = 4
    ERREUR_1304_NUMÉRO_APOSTROPHE_AVANT_NOM_ORDINATEUR = 7
    ERREUR_3035_NOMBRE_APOSTROPHES = 8
    ERREUR_3035_NUMÉRO_APOSTROPHE_AVANT_NOM_UTILISATEUR = 4
    ERREUR_3035_NUMÉRO_APOSTROPHE_AVANT_NOM_ORDINATEUR = 7
    ERREUR_3049_NOMBRE_APOSTROPHES = 8
    ERREUR_3049_NUMÉRO_APOSTROPHE_AVANT_NOM_UTILISATEUR = 4
    ERREUR_3049_NUMÉRO_APOSTROPHE_AVANT_NOM_ORDINATEUR = 7
    ERREUR_3200_NOMBRE_APOSTROPHES = 3
    ERREUR_3200_NUMÉRO_APOSTROPHE_AVANT_NOM_TABLE = 2
    ERREUR_3201_NOMBRE_APOSTROPHES = 3
    ERREUR_3201_NUMÉRO_APOSTROPHE_AVANT_NOM_TABLE = 2
    ERREUR_3260_NOMBRE_APOSTROPHES = 6
    ERREUR_3260_NUMÉRO_APOSTROPHE_AVANT_NOM_UTILISATEUR = 2
    ERREUR_3260_NUMÉRO_APOSTROPHE_AVANT_NOM_ORDINATEUR = 5
    ERREUR_3261_NOMBRE_APOSTROPHES = 8
    ERREUR_3261_NUMÉRO_APOSTROPHE_AVANT_NOM_UTILISATEUR = 4
    ERREUR_3261_NUMÉRO_APOSTROPHE_AVANT_NOM_ORDINATEUR = 7
End Enum

'   Note: Attention au choix du nom de la source de données système (mstrDSN).
'       Ce choix se fait via l'Administrateur de sources de données ODBC (onglet Source de données système), et il implémente les clés HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI et HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources.
'       Cependant, l'implémentation se fait parallèlement à la mise à jour du fichier C:\WINDOWS\ODBC.INI. Le problème est le suivant:
'           - Chaque création d'une source de données "utilisateur" ou "système" via l'Administrateur de sources de données ODBC provoque l'écriture d'une section correspondante dans le fichier C:\WINDOWS\ODBC.INI, sous réserve que cette section n'existe pas déjà (cas d'un même nom utilisé pour une source de données "utilisateur" et "système").
'           - L'implémentation de la clé HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI\<nom de la source de données> se fait parfois de manière aléatoire, sans qu'il soit possible de dégager une règle claire:
'               - le nom Adoptions utilisé comme source de données "système" (précédemment utilisé comme source de données "utilisateur") crée un dysfonctionnement: la solution consiste alors à utiliser un autre nom, puis à le renommer chaque fois qu'il apparaît sous la clé HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI, ainsi qu'à renommer la section du fichier C:\WINDOWS\ODBC.INI.
'               - en revanche, le nom AdoptionsBis utilisé à la fois comme source de données "utilisateur" et "système" ne crée pas de dysfonctionnement ...
'           - Par ailleurs, la suppression d'une source de données via l'Administrateur de sources de données ODBC provoque la suppression de la section correspondante dans le fichier C:\WINDOWS\ODBC.INI.
'       Une fois la source de données nommée via l'Administrateur de sources de données ODBC, il convient de donc vérifier l'adéquation entre le registre et le fichier C:\WINDOWS\ODBC.INI.
Private Const mstrDSN As String = "Adoptions", mstrDBQ As String = "DBQ", mstrDESCRIPTION As String = "Description", mstrDESCRIPTION_DÉFAUT As String = "Connexion ODBC pour la base Adoptions"
Private Const mstrREGISTRE_ODBC As String = "Software\ODBC\ODBC.INI", mstrREGISTRE_ODBC_ADOPTIONS As String = mstrREGISTRE_ODBC & "\" & mstrDSN, mstrREGISTRE_ODBC_DBQ As String = mstrREGISTRE_ODBC_ADOPTIONS & "\" & mstrDBQ, mstrREGISTRE_ODBC_DESCRIPTION As String = mstrREGISTRE_ODBC_ADOPTIONS & "\" & mstrDESCRIPTION

Private mdbeAdoptions As DBEngine
Private mwksODBCAdoptions As Workspace
Private mconnODBCAdoptions As Connection
Private mcolParamètresÉchéancier As Collection
Private mstrChaîneSQL As String, mstrNomBaseAvecChemin As String
Private mlngDernièreErreurDAO As DAOErreurs
Private mlngTypeRecordset As DAO.RecordsetTypeEnum


'Rafraîchir un jeu d'enregistrements
'   Les valeurs prises par mlngDernièreErreurDAO sont:  (*) daoENREGISTREMENT_VERROUILLÉ:   Les données sont verrouillées par une transaction
'                                                       (#) daoAUCUN_ENREGISTREMENT:        Le jeu ne contient pas d'enregistrement
'                                                       (*) daoAUCUNE:                      La lecture a retourné un jeu d'enregistrements
'       (*) la valeur est retournée et peut donc être interceptée par la procédure appelante
'       (#) la valeur n'est retournée que si aucun message n'est défini (message si recordset vide)
'   prstCible:          jeu d'enregistrements sur lequel porte la mise à jour ou la suppression
'   pbooGérerSablier:   booléen précisant s'il faut gérer le sablier
'   pstrListeMessages:  liste de messages déterminant les messages supplémentaires à afficher en cas d'erreur au format '<message si verrouillage[¤pictogramme]>[|<message si recordset vide>]'
'                       pour chaque message, 3 formats possibles:   "":                         aucun message n'est affiché
'                                                                   "<message>":                le message est affiché avec le pictogramme vbExclamation
'                                                                   "<message>¤<pictogramme>":  le message est affiché avec le pictogramme désigné

Public Sub ActualiserRecordset(prstCible As Recordset, pbooGérerSablier As Boolean, ByVal pstrListeMessages As String)
Const pkstrNomProcédure As String = "ActualiserRecordset"
Const pkintVerrouillage As Integer = 0, pkintRecordsetVide As Integer = 1
Dim pdatHeureRéférence As Date
Dim pstrListesMessages() As String, pstrMessages() As String
Dim pbooVerrouillageActivé As Boolean
On Error GoTo Err_ActualiserRecordset
    
    If pbooGérerSablier Then Screen.MousePointer = vbHourglass
    
    pstrListesMessages() = Split(pstrListeMessages, gstrSÉPARATEUR_1)
    While Not UBound(pstrListesMessages) = pkintRecordsetVide
        pstrListeMessages = pstrListeMessages & gstrSÉPARATEUR_1
        pstrListesMessages() = Split(pstrListeMessages, gstrSÉPARATEUR_1)
    Wend
    
    pdatHeureRéférence = Now
    
    With prstCible
    
        Do
        
'           Si les données ont été verrouillées par une transaction
            If pbooVerrouillageActivé Then
'               Avertir l'utilisateur du verrouillage au cas où le délai maximal avant l'affichage d'un message serait atteint
                GérerVerrouillage pbooVerrouillageActivé, pdatHeureRéférence
'               Annuler l'opération si l'utilisateur le demande
                If mlngDernièreErreurDAO = daoENREGISTREMENT_VERROUILLÉ Then
'                   Si un message supplémentaire doit être affiché
                    If Not pstrListesMessages(pkintVerrouillage) = "" Then
                        pstrMessages() = Split(pstrListesMessages(pkintVerrouillage), gstrSÉPARATEUR_2)
'                       Si aucun pictogramme n'est défini
                        If UBound(pstrMessages) = 0 Then
                            ReDim Preserve pstrMessages(1)
                            pstrMessages(1) = vbExclamation
                        End If
                        Err.Raise mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION
                    End If
                    If pbooGérerSablier Then Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            
            mlngDernièreErreurDAO = daoAUCUNE
    
'           Rafraîchir le jeu d'enregistrements
            .Requery
            DoEvents
            
'       Boucler tant qu'une transaction verrouille les données
        Loop While pbooVerrouillageActivé
        
    End With
    
    RemplirRecordsetEnTotalité prstCible
    
'   Si le jeu ne contient pas d'enregistrement
    If mlngDernièreErreurDAO = daoAUCUN_ENREGISTREMENT Then
'       En informer l'utilisateur si cela est demandé
        If Not pstrListesMessages(pkintRecordsetVide) = "" Then Err.Raise mesRECORDSET_VIDE
        If pbooGérerSablier Then Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If pbooGérerSablier Then Screen.MousePointer = vbDefault
    
    Exit Sub

Err_ActualiserRecordset:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                                     pintTraitement = mesINTERROMPRE
        Case daoERREUR_03146:                                   If mdbeAdoptions.Errors(0).Number = daoERREUR_03261 Then pbooVerrouillageActivé = True: Resume Next
                                                                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
        Case mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrMessages(0), pstrMessages(1))
        Case mesRECORDSET_VIDE:                                 pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrListesMessages(pkintRecordsetVide), prstCible.Name)
        Case Else:                                              pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub


'Gérer une transaction
'   plngMéthode:    méthode souhaitée

Public Sub GérerTransaction(plngMéthode As DAOTransactions)
Const pkstrNomProcédure As String = "GérerTransaction"
Dim pobjTransaction As Object
Dim prstCible As Recordset
Dim pvarSignets() As Variant
Dim pintIndiceTableau As Integer
On Error GoTo Err_GérerTransaction
    
#If TablesEnLocalOuAttachées Then
    Set pobjTransaction = mwksODBCAdoptions
#Else
    Set pobjTransaction = mdbeAdoptions
#End If

    Select Case plngMéthode
    
'       Débuter une transaction
        Case daoOUVERTURE
        
            pobjTransaction.BeginTrans
        
'       Valider une transaction
        Case daoVALIDATION
        
            pobjTransaction.CommitTrans
        
'       Annuler une transaction
        Case daoANNULATION
        
            pintIndiceTableau = -1
        
'           Pour chaque recordset de la connexion ODBC
            For Each prstCible In mconnODBCAdoptions.Recordsets
                pintIndiceTableau = pintIndiceTableau + 1
                ReDim Preserve pvarSignets(pintIndiceTableau)
                If Not prstCible Is Nothing Then
                    With prstCible
#If TablesEnLocalOuAttachées Then
'                       Si le recordset n'est pas de type défilement en avant seulement, s'il compte un ou plusieurs enregistrements et qu'aucune suppression n'a été tentée sur l'enregistrement courant, implémenter le tableau des signets par sa position courante
                        If Not (.Type = dbOpenForwardOnly Or .RecordCount = 0 Or .RecordStatus = dbRecordDeleted) Then pvarSignets(pintIndiceTableau) = .Bookmark
#Else
'                       Si le recordset n'est pas de type défilement en avant seulement et s'il compte un ou plusieurs enregistrements, implémenter le tableau des signets par sa position courante
                        If Not (.Type = dbOpenForwardOnly Or .RecordCount = 0) Then pvarSignets(pintIndiceTableau) = .Bookmark
#End If
                    End With
                End If
            Next prstCible
    
'           Annuler la transaction
            pobjTransaction.Rollback
            
            pintIndiceTableau = -1
        
'           Pour chaque recordset de la connexion ODBC
            For Each prstCible In mconnODBCAdoptions.Recordsets
                pintIndiceTableau = pintIndiceTableau + 1
'               Si un signet a été défini, restaurer la position courante du recordset
                If Not IsEmpty(pvarSignets(pintIndiceTableau)) Then prstCible.Bookmark = pvarSignets(pintIndiceTableau)
            Next prstCible
    
    End Select
    
    Exit Sub
    
Err_GérerTransaction:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub


'Vérifier s'il faut avertir l'utilisateur d'un verrouillage, et si oui, lui proposer de poursuivre ou d'abandonner l'opération
'   pbooVerrouillageActivé: booléen précisant l'état du verrouillage
'   pdatHeureRéférence:     heure de la première tentative de mise à jour

Private Sub GérerVerrouillage(ByRef pbooVerrouillageActivé As Boolean, ByRef pdatHeureRéférence As Date)
Const pkstrNomProcédure As String = "GérerVerrouillage"
Dim pstrDescriptionErreurODBC As String, pstrNomUtilisateur As String, pstrNomOrdinateur As String, pstrNomTable As String
Dim pintPositionApostrophe() As Integer, pintNuméroApostrophe() As Integer, pintIndiceTableau As Integer, pintPositionDépartRecherche As Integer
On Error GoTo Err_GérerVerrouillage

    Select Case mdbeAdoptions.Errors(0).Number
'       S1000: [Microsoft][Pilote ODBC Microsoft Access] Impossible de verrouiller la table '<#1>'; actuellement utilisée par l'utilisateur '<#2>' sur l'ordinateur '<#3>'.
        Case daoERREUR_01304: ReDim pintPositionApostrophe(1 To ERREUR_1304_NOMBRE_APOSTROPHES) As Integer: ReDim pintNuméroApostrophe(0 To 1) As Integer: pintNuméroApostrophe(0) = CInt(ERREUR_1304_NUMÉRO_APOSTROPHE_AVANT_NOM_UTILISATEUR): pintNuméroApostrophe(1) = CInt(ERREUR_1304_NUMÉRO_APOSTROPHE_AVANT_NOM_ORDINATEUR)
'       S1000: [Microsoft][Pilote ODBC Microsoft Access] Impossible de verrouiller la table '<#1>'; actuellement utilisée par l'utilisateur '<#2>' sur l'ordinateur '<#3>'.
        Case daoERREUR_03035: ReDim pintPositionApostrophe(1 To ERREUR_3035_NOMBRE_APOSTROPHES) As Integer: ReDim pintNuméroApostrophe(0 To 1) As Integer: pintNuméroApostrophe(0) = CInt(ERREUR_3035_NUMÉRO_APOSTROPHE_AVANT_NOM_UTILISATEUR): pintNuméroApostrophe(1) = CInt(ERREUR_3035_NUMÉRO_APOSTROPHE_AVANT_NOM_ORDINATEUR)
'       S1000: [Microsoft][Pilote ODBC Microsoft Access] Impossible de verrouiller la table '<#1>'; actuellement utilisée par l'utilisateur '<#2>' sur l'ordinateur '<#3>'.
        Case daoERREUR_03049: ReDim pintPositionApostrophe(1 To ERREUR_3049_NOMBRE_APOSTROPHES) As Integer: ReDim pintNuméroApostrophe(0 To 1) As Integer: pintNuméroApostrophe(0) = CInt(ERREUR_3049_NUMÉRO_APOSTROPHE_AVANT_NOM_UTILISATEUR): pintNuméroApostrophe(1) = CInt(ERREUR_3049_NUMÉRO_APOSTROPHE_AVANT_NOM_ORDINATEUR)
'       S1000: [Microsoft][Pilote ODBC Microsoft Access] Impossible de mettre à jour; actuellement verrouillé par l'utilisateur '<#1>' sur l'ordinateur '<#2>'.
        Case daoERREUR_03260: ReDim pintPositionApostrophe(1 To ERREUR_3260_NOMBRE_APOSTROPHES) As Integer: ReDim pintNuméroApostrophe(0 To 1) As Integer: pintNuméroApostrophe(0) = CInt(ERREUR_3260_NUMÉRO_APOSTROPHE_AVANT_NOM_UTILISATEUR): pintNuméroApostrophe(1) = CInt(ERREUR_3260_NUMÉRO_APOSTROPHE_AVANT_NOM_ORDINATEUR)
'       S1000: [Microsoft][Pilote ODBC Microsoft Access] La table '<#1>' est verrouillée exclusivement par l'utilisateur '<#2>' sur l'ordinateur '<#3>'.
        Case daoERREUR_03261: ReDim pintPositionApostrophe(1 To ERREUR_3261_NOMBRE_APOSTROPHES) As Integer: ReDim pintNuméroApostrophe(0 To 1) As Integer: pintNuméroApostrophe(0) = CInt(ERREUR_3261_NUMÉRO_APOSTROPHE_AVANT_NOM_UTILISATEUR): pintNuméroApostrophe(1) = CInt(ERREUR_3261_NUMÉRO_APOSTROPHE_AVANT_NOM_ORDINATEUR)
    End Select
    
'   Si le délai maximal avant l'affichage d'un message informant du verrouillage des données est atteint
    If DateDiff("s", pdatHeureRéférence, Now) > 5 Then
    
        pstrDescriptionErreurODBC = mdbeAdoptions.Errors(0).Description
    
        Select Case mdbeAdoptions.Errors(0).Number
            Case daoERREUR_01304, daoERREUR_03035, daoERREUR_03049, daoERREUR_03260, daoERREUR_03261
'               Extraire les positions des apostrophes dans le message
                For pintIndiceTableau = 1 To UBound(pintPositionApostrophe)
                    If pintIndiceTableau = 1 Then pintPositionDépartRecherche = 1 Else pintPositionDépartRecherche = pintPositionApostrophe(pintIndiceTableau - 1) + 1
                    pintPositionApostrophe(pintIndiceTableau) = InStr(pintPositionDépartRecherche, pstrDescriptionErreurODBC, "'")
                Next pintIndiceTableau
        End Select

        Select Case mdbeAdoptions.Errors(0).Number
#If Not TablesEnLocalOuAttachées Then
            Case daoERREUR_01809
'               Proposer la poursuite de l'opération
                Select Case modMessages.lngRéponseUtilisateur(mesQUESTION_001)
'                   L'utilisateur décide la poursuite de l'opération
                    Case vbYes
                        pdatHeureRéférence = Now
'                   L'utilisateur décide l'abandon de l'opération
                    Case vbNo
                        mlngDernièreErreurDAO = daoENREGISTREMENT_VERROUILLÉ
                End Select
#End If
            Case daoERREUR_01304, daoERREUR_03035, daoERREUR_03049, daoERREUR_03260, daoERREUR_03261
'               Extraire du message d'erreur DAO les noms de l'utilisateur et de l'ordinateur à l'origine du verrouillage
                pstrNomUtilisateur = Mid(pstrDescriptionErreurODBC, pintPositionApostrophe(pintNuméroApostrophe(0)) + 1, pintPositionApostrophe(pintNuméroApostrophe(0) + 1) - pintPositionApostrophe(pintNuméroApostrophe(0)) - 1)
                pstrNomOrdinateur = Mid(pstrDescriptionErreurODBC, pintPositionApostrophe(pintNuméroApostrophe(1)) + 1, pintPositionApostrophe(pintNuméroApostrophe(1) + 1) - pintPositionApostrophe(pintNuméroApostrophe(1)) - 1)
'               Proposer la poursuite de l'opération
                Select Case modMessages.lngRéponseUtilisateur(mesQUESTION_001, pstrNomUtilisateur, pstrNomOrdinateur)
'                   L'utilisateur décide la poursuite de l'opération
                    Case vbYes
                        pdatHeureRéférence = Now
'                   L'utilisateur décide l'abandon de l'opération
                    Case vbNo
                        mlngDernièreErreurDAO = daoENREGISTREMENT_VERROUILLÉ
                End Select
        End Select
        
    End If
    
    pbooVerrouillageActivé = False
    
    Exit Sub
    
Err_GérerVerrouillage:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub


'Implémenter les collections
'   pstrMessageSiVerrouillage:  message à afficher si les données sont verrouillées par une transaction
'   pvarArgument():             tableau de type ParamArray constitué de la liste des collections à implémenter

Public Sub ImplémenterCollections(pstrMessageSiVerrouillage As String, ParamArray pvarArgument() As Variant)
Const pkstrNomProcédure As String = "ImplémenterCollections"
Dim prstCollection As Recordset
Dim plngCode As Long
Dim pintIndiceArgument As Integer
Dim pbooCourAdministrativeAppelParamétrée As Boolean
On Error GoTo Err_ImplémenterCollections

'   Pour chaque argument passé en référence
    For pintIndiceArgument = LBound(pvarArgument) To UBound(pvarArgument)
    
        Select Case pvarArgument(pintIndiceArgument)
            
'           Paramètres de l'échéancier
            Case "Paramètres de l'échéancier"

                Set mcolParamètresÉchéancier = Nothing: Set mcolParamètresÉchéancier = New Collection
                
'               Lire les paramètres
                Set prstCollection = modDAO.rstSQLExtraction(daoEXTRACTION_001, True, pstrMessageSiVerrouillage)
                If Not (mlngDernièreErreurDAO = daoAUCUNE Or mlngDernièreErreurDAO = daoAUCUN_ENREGISTREMENT) Then Exit Sub

                With prstCollection
                
'                   Pour chaque paramètre
                    While Not .EOF
                    
'                       Implémenter la collection
'                           Clé: code paramètre, Élément: nombre de mois
'                           Ex: 2, 8
                        mcolParamètresÉchéancier.Add CInt(![PARA_NBMOIS]), CStr(![PARA_CODE])
        
'                       Lire le suivant
                        .MoveNext
                        
                    Wend
            
                End With
                
'           Recours contentieux
            Case "Recours contentieux"

                Set gclsProvenances = Nothing: Set gclsProvenances = New clsCodesLibellés
                Set gclsRecoursContentieux = Nothing: Set gclsRecoursContentieux = New clsRecoursContentieux: gclsRecoursContentieux.lngTribunalAdministratif = comINDÉFINI: gclsRecoursContentieux.lngCourAdministrativeAppel = comINDÉFINI: gclsRecoursContentieux.lngConseilÉtat = comINDÉFINI: Set gclsRecoursContentieux.clsRecoursContentieux = New clsCodesLibellés
                
                plngCode = AdoptionsProvenances.adpDEMANDE_INITIALE - 1
                plngCode = plngCode + 1: gclsProvenances.Ajout plngCode, "Demande initiale", pvarClé:=CStr(plngCode)
                plngCode = plngCode + 1: gclsProvenances.Ajout plngCode, "Recours administratif", pvarClé:=CStr(plngCode)
                
'               Lire les types
                Set prstCollection = modDAO.rstSQLExtraction(daoEXTRACTION_012, True, pstrMessageSiVerrouillage)
                If Not (mlngDernièreErreurDAO = daoAUCUNE Or mlngDernièreErreurDAO = daoAUCUN_ENREGISTREMENT) Then Exit Sub

                With prstCollection

'                   Pour chaque type
                    While Not .EOF

'                       Implémenter la collection des provenances d'agrément
                        gclsProvenances.Ajout ![TYRE_NUM], ![TYRE_NOM], pvarClé:=CStr(![TYRE_NUM])
                        
'                       Implémenter la collection des types de recours contentieux
                        gclsRecoursContentieux.clsRecoursContentieux.Ajout ![TYRE_NUM], ![TYRE_NOM], pvarClé:=CStr(![TYRE_NUM])
                        gclsRecoursContentieux.lngPlusHautRecours = ![TYRE_NUM]
                        
                        Select Case ![TYRE_NOM]
                            Case gclsRecoursContentieux.strTribunalAdministratif:   gclsRecoursContentieux.lngTribunalAdministratif = ![TYRE_NUM]
                            Case gclsRecoursContentieux.strCourAdministrativeAppel: gclsRecoursContentieux.lngCourAdministrativeAppel = ![TYRE_NUM]
                            Case gclsRecoursContentieux.strConseilÉtat:             gclsRecoursContentieux.lngConseilÉtat = ![TYRE_NUM]
                        End Select
                        
'                       Lire le suivant
                        .MoveNext

                    Wend

                End With
                
        End Select
    
    Next pintIndiceArgument
    
    Exit Sub
    
Err_ImplémenterCollections:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub


'Retourner la valeur d'un paramètre de l'échéancier
'   plngPARA_CODE:  code paramètre correspondant à la colonne PARA.PARA_CODE

Public Property Get intParamètreÉchéancier(plngPARA_CODE As DAOParamètresÉchéancier) As Integer
Const pkstrNomProcédure As String = "intParamètreÉchéancier"
Dim pstrClé As String
Dim pbooÉlémentManquant As Boolean
On Error GoTo Err_intParamètreÉchéancier

    pstrClé = CStr(plngPARA_CODE)
    intParamètreÉchéancier = mcolParamètresÉchéancier(pstrClé)
    If pbooÉlémentManquant Then Err.Raise mesENREGISTREMENT_MANQUANT
    
    Exit Property
    
Err_intParamètreÉchéancier:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                 pintTraitement = mesINTERROMPRE
        Case comERREUR_00005:               pbooÉlémentManquant = True: Resume Next
        Case mesENREGISTREMENT_MANQUANT:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, "PARA", CStr(plngPARA_CODE))
        Case Else:                          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Property


'Retourner le résultat de la dernière opération DAO

Public Property Get lngDernièreErreurDAO() As DAOErreurs
Const pkstrNomProcédure As String = "lngDernièreErreurDAO"
On Error GoTo Err_lngDernièreErreurDAO

    lngDernièreErreurDAO = mlngDernièreErreurDAO
    
    Exit Property
    
Err_lngDernièreErreurDAO:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Property


'Retourner le nombre d'enregistrements concernés par la dernière requête de suppression, de mise à jour ou d'insertion

Public Property Get lngNombreEnregistrementsAffectés() As Long
Const pkstrNomProcédure As String = "lngNombreEnregistrementsAffectés"
On Error GoTo Err_lngNombreEnregistrementsAffectés

    If mconnODBCAdoptions Is Nothing Then Exit Property
    
    lngNombreEnregistrementsAffectés = mconnODBCAdoptions.RecordsAffected
    
    Exit Property
    
Err_lngNombreEnregistrementsAffectés:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Property


'Ouvrir la connexion ODBC

Public Sub OuvrirConnexion()
Const pkstrNomProcédure As String = "OuvrirConnexion"
Dim pstrClé As String, pstrDescription As String, pstrErreur As String
Dim plngRetour As Long, pbooÉcrireRegistre As Long
Dim pintAttributs As Integer
Dim pbooErreurPossible As Boolean, pbooBaseInaccessible As Boolean
On Error GoTo Err_OuvrirConnexion

'   Si la connexion est déjà ouverte, sortir
    If Not (mdbeAdoptions Is Nothing Or mwksODBCAdoptions Is Nothing Or mconnODBCAdoptions Is Nothing) Then Exit Sub
    
    Screen.MousePointer = vbHourglass

'   Vérifier l'existence de la clé 'ODBC.INI' dans le registre
    pbooÉcrireRegistre = False: pstrClé = mstrREGISTRE_ODBC
    plngRetour = modCommun.lngVérificationExistenceCléRegistre(apiHKEY_LOCAL_MACHINE, mstrREGISTRE_ODBC)
    If Not plngRetour = apiERROR_SUCCESS Then Err.Raise mesREGISTRE_LECTURE_ÉCRITURE
    
'   Vérifier l'existence du driver ODBC pour la base Adoptions
    plngRetour = modCommun.lngVérificationExistenceCléRegistre(apiHKEY_LOCAL_MACHINE, mstrREGISTRE_ODBC_ADOPTIONS)
    If Not plngRetour = apiERROR_SUCCESS Then Err.Raise mesMESSAGE_001
    
'   Vérifier la localisation de la base Adoptions
    pbooÉcrireRegistre = False: pstrClé = mstrREGISTRE_ODBC_DBQ
    plngRetour = modCommun.lngLectureValeurRegistre(apiHKEY_LOCAL_MACHINE, mstrREGISTRE_ODBC_ADOPTIONS, mstrDBQ, apiREG_SZ, mstrNomBaseAvecChemin)
    If Not plngRetour = apiERROR_SUCCESS Then Err.Raise mesREGISTRE_LECTURE_ÉCRITURE
    
'   Vérifier sa disponibilité
    pbooErreurPossible = True
    pintAttributs = GetAttr(mstrNomBaseAvecChemin)
    If pbooBaseInaccessible Then Err.Raise mesMESSAGE_022
    pbooErreurPossible = False
    
'   Lire sa description
    pbooÉcrireRegistre = True: pstrClé = mstrREGISTRE_ODBC_DESCRIPTION
    plngRetour = modCommun.lngLectureValeurRegistre(apiHKEY_LOCAL_MACHINE, mstrREGISTRE_ODBC_ADOPTIONS, mstrDESCRIPTION, apiREG_SZ, pstrDescription, True, mstrDESCRIPTION_DÉFAUT)
    If Not plngRetour = apiERROR_SUCCESS Then Err.Raise mesREGISTRE_LECTURE_ÉCRITURE
    
'   Ouvrir la connexion ODBC
    Set mdbeAdoptions = New DBEngine
    Set mwksODBCAdoptions = mdbeAdoptions.CreateWorkspace("Espace ODBCDirect - " & Now, "", "", dbUseODBC)
#If TablesEnLocalOuAttachées Then
    mwksODBCAdoptions.DefaultCursorDriver = dbUseClientBatchCursor
#End If
    Set mconnODBCAdoptions = mwksODBCAdoptions.OpenConnection(pstrDescription, , , "ODBC;DSN=" & mstrDSN)
    
    Screen.MousePointer = vbDefault

    Exit Sub

Err_OuvrirConnexion:
Dim pintTraitement As MessagesActions
    If pbooErreurPossible And (Err.Number = comERREUR_00053 Or Err.Number = comERREUR_00068 Or Err.Number = comERREUR_00071 Or Err.Number = comERREUR_00076) Then pbooBaseInaccessible = True: pstrErreur = Format(Err.Number) & " - " & Err.Description: Resume Next
    Select Case Err.Number
        Case vbObjectError:                 pintTraitement = mesINTERROMPRE
        Case mesREGISTRE_LECTURE_ÉCRITURE:  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, plngRetour, IIf(pbooÉcrireRegistre, "d'écriture", "de lecture"), pstrClé)
        Case mesMESSAGE_001:                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
        Case mesMESSAGE_022:                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, "La base définie par le driver ODBC", "e", mstrNomBaseAvecChemin, pstrErreur)
        Case Else:                          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Remplir l'intégralité d'un jeu d'enregistrements
'   Permet de connaître le nombre exact d'enregistrements et, au cas où une nouvelle transaction serait ouverte, d'éviter l'erreur ODBC n°72 ('Attribute cannot be set now')
'   Positionne mlngDernièreErreurDAO à la valeur daoAUCUN_ENREGISTREMENT si le jeu ne contient pas d'eregistrement
'   prstSource:                                 jeu d'enregistrements dont on souhaite le remplissage
'   pbooSePositionnerSurPremierEnregistrement:  booléen précisant si le premier enregistrement doit devenir l'enregistrement courant [facultatif: s'il est manquant, le premier enregistrement deviendra l'enregistrement courant]

Private Sub RemplirRecordsetEnTotalité(prstSource As Recordset, Optional pbooSePositionnerSurPremierEnregistrement As Boolean = True)
Const pkstrNomProcédure As String = "RemplirRecordsetEnTotalité"
On Error GoTo Err_RemplirRecordsetEnTotalité
    
    With prstSource
    
'       Si le jeu d'enregistrements est de type défilement en avant seulement, sortir
        If .Type = dbOpenForwardOnly Then Exit Sub
    
'       Si le jeu ne contient pas d'enregistrement
        If .RecordCount = 0 Then
            mlngDernièreErreurDAO = daoAUCUN_ENREGISTREMENT
            Exit Sub
        End If
        
'       Atteindre le dernier enregistrement
        .MoveLast
'       Atteindre le premier si demandé
        If pbooSePositionnerSurPremierEnregistrement Then .MoveFirst
        
    End With
    
    Exit Sub
    
Err_RemplirRecordsetEnTotalité:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Retourner un jeu d'enregistrements
'   Les valeurs prises par mlngDernièreErreurDAO sont:  (*) daoENREGISTREMENT_VERROUILLÉ:   Les données sont verrouillées par une transaction
'                                                       (#) daoAUCUN_ENREGISTREMENT:        Le jeu ne contient pas d'enregistrement
'                                                       (*) daoAUCUNE:                      La lecture a retourné un jeu d'enregistrements
'       (*) la valeur est retournée et peut donc être interceptée par la procédure appelante
'       (#) la valeur n'est retournée que si aucun message n'est défini (message si recordset vide)
'   plngNuméroRequête:  numéro de la requête
'   pbooGérerSablier:   booléen précisant s'il faut gérer le sablier
'   pstrListeMessages:  liste de messages déterminant les messages supplémentaires à afficher en cas d'erreur au format '<message si verrouillage[¤pictogramme]>[|<message si recordset vide>]'
'                       pour chaque message, 3 formats possibles:   "":                         aucun message n'est affiché
'                                                                   "<message>":                le message est affiché avec le pictogramme vbExclamation
'                                                                   "<message>¤<pictogramme>":  le message est affiché avec le pictogramme désigné
'   pvarArgument():     tableau d'arguments

Public Function rstSQLExtraction(plngNuméroRequête As DAOExtractions, pbooGérerSablier As Boolean, ByVal pstrListeMessages As String, ParamArray pvarArgument() As Variant) As Recordset
Const pkstrNomProcédure As String = "rstSQLExtraction"
Const pkintVerrouillage As Integer = 0, pkintRecordsetVide As Integer = 1
Dim plngTypeRecordset As DAO.RecordsetTypeEnum
Dim pdatHeureRéférence As Date
Dim pstrChaîneSQL As String, pstrArguments() As String, pstrListesMessages() As String, pstrMessages() As String
Dim pintIndiceArgument As Integer
Dim pbooVerrouillageActivé As Boolean
On Error GoTo Err_rstSQLExtraction

    If pbooGérerSablier Then Screen.MousePointer = vbHourglass
    
    pstrListesMessages() = Split(pstrListeMessages, gstrSÉPARATEUR_1)
    While Not UBound(pstrListesMessages) = pkintRecordsetVide
        pstrListeMessages = pstrListeMessages & gstrSÉPARATEUR_1
        pstrListesMessages() = Split(pstrListeMessages, gstrSÉPARATEUR_1)
    Wend
    
    If Not UBound(pvarArgument) = comINDÉFINI Then ReDim pstrArguments(LBound(pvarArgument) To UBound(pvarArgument))
    
'   Pour chaque argument passé en référence
    For pintIndiceArgument = LBound(pvarArgument) To UBound(pvarArgument)
        Select Case VarType(pvarArgument(pintIndiceArgument))
            Case vbBoolean
                pstrArguments(pintIndiceArgument) = "= " & IIf(pvarArgument(pintIndiceArgument), "TRUE", "FALSE")
            Case vbDate
                pstrArguments(pintIndiceArgument) = "= " & strSQLDateFormatée(pvarArgument(pintIndiceArgument))
            Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                pstrArguments(pintIndiceArgument) = "= " & strSQLNombreFormaté(pvarArgument(pintIndiceArgument))
            Case vbNull
                pstrArguments(pintIndiceArgument) = "IS NULL"
            Case vbString
                pstrArguments(pintIndiceArgument) = "= '" & strSQLChaîneFormatée(pvarArgument(pintIndiceArgument)) & "'"
        End Select
    Next pintIndiceArgument

    Select Case plngNuméroRequête

'       Extraire un individu
'           INDI_NOM|INDI_PREN
'           Ex: RIVOAL|Didier
        Case daoINDIVIDU
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "INDI.INDI_NOM, INDI.INDI_PREN " _
                          & "FROM INDI " _
                          & "WHERE (INDI.DOSS_NUM " & pstrArguments(0) & " AND INDI.INDI_CDFR " & pstrArguments(1) & ")"
                          
'       Compter le nombre d'enregistrements
'           Nombre
'           Ex: 1
        Case daoNOMBRE
            plngTypeRecordset = dbOpenForwardOnly
            pstrChaîneSQL = "SELECT " _
                          & "COUNT(*) AS Nombre " _
                          & "FROM " & pvarArgument(0) & " " _
                          & "WHERE (" & pvarArgument(0) & "." & pvarArgument(1) & " " & pstrArguments(2) & ")"

'       Extraire la valeur maximale d'une colonne
'           Maximum
'           Ex: 3
        Case daoMAXIMUM
            plngTypeRecordset = dbOpenForwardOnly
            pstrChaîneSQL = "SELECT " _
                          & "MAX(" & pvarArgument(0) & "." & pvarArgument(1) & ") AS Maximum " _
                          & "FROM " & pvarArgument(0)
                          
'       Extraire les recours contentieux
'           TYRE_NUM|RECO_DDREC|RECO_DPREDEC|RECO_DDECI|RECO_DECI|RECO_REFUS|RECO_dtmAgrementDelivrance|DEMA_NUM|u02_dtmHistoCommissionReunion
'           Ex: 1|10/09/2003|10/06/2004|27/01/2005|R|décision du PCG fondée||9907903|04/04/2003
        Case daoRECOURS_CONTENTIEUX
            plngTypeRecordset = pvarArgument(0)
            pstrChaîneSQL = "SELECT * " _
                          & "FROM RECO " _
                          & "WHERE (RECO.DEMA_NUM " & pstrArguments(1) & " AND RECO.u02_dtmHistoCommissionReunion " & pstrArguments(2) & ") " _
                          & "ORDER BY RECO.TYRE_NUM DESC"
    
'       Extraire l'historique des psychologues/assistantes sociales
        Case daoPSYCHOLOGUES__OU__ASSISTANTES_SOCIALES
            plngTypeRecordset = pvarArgument(0)
            If plngTypeRecordset = dbOpenSnapshot Then
'               DEMA_NUM|u01_intHistoIntervenantCadre|u01_booHistoIntervenantPsychologue|u01_dtmHistoIntervenantDemande|u01_intHistoIntervenantNumero|u01_dtmHistoIntervenantPrevu|u01_dtmHistoIntervenantRecu|TRSO_NUM|TRSO_NOM
'               Ex: 9908578|0|Non|06/03/2008|1|25/10/2008||176|COLLIN
                pstrChaîneSQL = "SELECT " _
                              & "u01_HistoIntervenants.*, TRSO.TRSO_NOM " _
                              & "FROM TRSO RIGHT JOIN u01_HistoIntervenants ON TRSO.TRSO_NUM = u01_HistoIntervenants.TRSO_NUM " _
                              & "WHERE (u01_HistoIntervenants.DEMA_NUM " & pstrArguments(1) & " AND u01_HistoIntervenants.u01_booHistoIntervenantPsychologue " & pstrArguments(2) & ") " _
                              & "ORDER BY u01_HistoIntervenants.u01_dtmHistoIntervenantDemande DESC, u01_HistoIntervenants.u01_intHistoIntervenantNumero DESC"
            Else
'               DEMA_NUM|u01_intHistoIntervenantCadre|u01_booHistoIntervenantPsychologue|u01_dtmHistoIntervenantDemande|u01_intHistoIntervenantNumero|u01_dtmHistoIntervenantPrevu|u01_dtmHistoIntervenantRecu|TRSO_NUM
'               Ex: 9908578|0|Non|06/03/2008|1|25/10/2008||176
                pstrChaîneSQL = "SELECT * " _
                              & "FROM u01_HistoIntervenants " _
                              & "WHERE (u01_HistoIntervenants.DEMA_NUM " & pstrArguments(1) & " AND u01_HistoIntervenants.u01_booHistoIntervenantPsychologue " & pstrArguments(2) & ") " _
                              & "ORDER BY u01_HistoIntervenants.u01_dtmHistoIntervenantDemande DESC, u01_HistoIntervenants.u01_intHistoIntervenantNumero DESC"
            End If
    
'       Extraire les dates de dépôt et de référence de la demande, et la date de réception du dossier
'           DEMA_DATE|DEMA_DREDOS|DEMA_DREF
'           Ex: 01/02/2008|27/02/2008|27/02/2008
        Case daoDEMANDE_DATES
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "DEMA.DEMA_DATE, DEMA.DEMA_DREDOS, DEMA.DEMA_DREF " _
                          & "FROM DEMA " _
                          & "WHERE (DEMA.DEMA_NUM " & pstrArguments(0) & ")"
    
'       Extraire le recours administratif
'           TRSO_NUMAS|TRSO_NUMPSY|READ_DDREC|READ_DCOM|READ_DPREDEC|READ_DDECI|READ_DECIC|READ_DECIF|READ_DDEMAS|READ_DPRRECAS|READ_RECAS|READ_DDEMPSY|READ_DPRRECPSY|READ_RECPSY|READ_REFUS|DEMA_NUM|u02_dtmHistoCommissionReunion
'           Ex: 194|7|10/01/2000|08/09/2000|10/10/2000|12/09/2000|A|A|14/01/2000|10/08/2000|25/07/2000|14/01/2000|10/08/2000|22/06/2000||9004003|08/09/2000
        Case daoRECOURS_ADMINISTRATIF
            plngTypeRecordset = pvarArgument(0)
            pstrChaîneSQL = "SELECT * " _
                          & "FROM READ " _
                          & "WHERE (READ.DEMA_NUM " & pstrArguments(1) & " AND READ.u02_dtmHistoCommissionReunion " & pstrArguments(2) & ")"
    
'       Extraire les surveillances d'un placement
'           TRSO_NUM_AssistanteSociale|PLAC_NUM|SUPL_DDSUR|SUPL_dtmAssistanteSocialeIntervention|SUPL_dtmAssistanteSocialeRecu|SUPL_DENVTGI|SUPL_dtmAssistanteSocialePrevu|TRSO_NUM_Psychologue|SUPL_dtmPsychologueIntervention|SUPL_dtmPsychologueRecu|SUPL_dtmPsychologuePrevu|SUPL_intNumero
'           Ex: 209|541|03/10/2001|27/12/2001|27/12/2001|21/02/2002|12/11/2001|209|27/12/2001|27/12/2001|12/11/2001|2
        Case daoSURVEILLANCES
            plngTypeRecordset = pvarArgument(0)
            pstrChaîneSQL = "SELECT * " _
                          & "FROM SUPL " _
                          & "WHERE (SUPL.PLAC_NUM " & pstrArguments(1) & ") " _
                          & "ORDER BY SUPL.SUPL_DDSUR DESC, SUPL.SUPL_intNumero DESC"
    
'       Extraire l'ensemble des décisions relatives à une demande
'           CommissionInitiale|Decision|DateCommission|DateSignature|Provenance|DateDepot
'           Ex: 01/07/2008|A|20/07/2008||-1|15/07/2008
        Case daoDÉCISIONS
            plngTypeRecordset = dbOpenSnapshot
'           Note: Laisser impérativement la lecture de la table READ en entête des deux autres de façon à:
'                   - définir les colonnes 'DateCommission', 'DateSignature' et 'DateDepot' comme étant de type dbDate (l'expression <NULL AS> en entête empêche ce typage et pose problème en cas d'affectation de la colonne par la suite)
'                   - définir la colonne 'Decision' comme étant de type dbText (la table u02_HistoCommissions proposant un entier pour cette colonne)
            pstrChaîneSQL = "SELECT " _
                          & "READ.u02_dtmHistoCommissionReunion AS CommissionInitiale, READ.READ_DECIF AS Decision, READ.READ_DCOM AS DateCommission, READ.READ_DDECI AS DateSignature, " & CStr(AdoptionsProvenances.adpRECOURS_ADMINISTRATIF) & " AS Provenance, READ_DDREC As DateDepot " _
                          & "FROM READ " _
                          & "WHERE (READ.DEMA_NUM " & pstrArguments(0) & ") " _
                          & "UNION SELECT " _
                          & "u02_HistoCommissions.u02_dtmHistoCommissionReunion AS CommissionInitiale, u02_HistoCommissions.u02_intHistoCommissionDecision AS Decision, u02_HistoCommissions.u02_dtmHistoCommissionReunion AS DateCommission, u02_HistoCommissions.u02_dtmHistoCommissionSignature AS DateSignature, " & CStr(AdoptionsProvenances.adpDEMANDE_INITIALE) & " AS Provenance, NULL AS DateDepot " _
                          & "FROM u02_HistoCommissions " _
                          & "WHERE (u02_HistoCommissions.DEMA_NUM " & pstrArguments(0) & ") " _
                          & "UNION SELECT " _
                          & "RECO.u02_dtmHistoCommissionReunion AS CommissionInitiale, RECO.RECO_DECI AS Decision, NULL AS DateCommission, RECO.RECO_dtmAgrementDelivrance AS DateSignature, RECO.TYRE_NUM AS Provenance, RECO_DDREC AS Depot " _
                          & "FROM RECO " _
                          & "WHERE (RECO.DEMA_NUM " & pstrArguments(0) & ") " _
                          & "ORDER BY CommissionInitiale ASC, Provenance ASC"
    
'       Extraire les enfants (souhaités ou accordés)
'           ENDE_NUM|ENDE_SEXE|ENDE_HAND|ENDE_TRAGE
'           Ex: 5108|I|N|00-01
        Case daoENFANTS_DEMANDÉS_OU_ACCORDÉS
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "ENDE.* " _
                          & "FROM ENDE, " & pvarArgument(0) & " " _
                          & "WHERE (ENDE.ENDE_NUM = " & pvarArgument(0) & ".ENDE_NUM AND " & pvarArgument(0) & ".DEMA_NUM " & pstrArguments(1) & ")"

'       Extraire les références du placement
'           DEMA_NUM|DOSS_NUM|DEMA_DATE|DEMA_DREDOS
'           Ex: 9908581|2008015|07/03/2002|30/03/2002
        Case daoPLACEMENT_RÉFÉRENCES
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "PLAC.DEMA_NUM, " _
                          & "DEMA.DOSS_NUM, DEMA.DEMA_DATE, DEMA.DEMA_DREDOS " _
                          & "FROM DEMA INNER JOIN PLAC ON DEMA.DEMA_NUM = PLAC.DEMA_NUM " _
                          & "WHERE (PLAC.PLAC_NUM " & pstrArguments(0) & ")"

'       modDAO: Extraire les paramètres de l'échéancier
'           PARA_CODE|PARA_NBMOIS
'           Ex: 2|8
        Case daoEXTRACTION_001
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "PARA.PARA_CODE, PARA.PARA_NBMOIS " _
                          & "FROM PARA"
                          
'       frm_DecisionCommission: Extraire la dernière/prochaine commission
'           DEMA_DATE|DEMA_CFRAT|DEMA_DREF|u02_dtmHistoCommissionReunion|u02_dtmHistoCommissionSignature|u02_dtmHistoCommissionPrevu|u02_intHistoCommissionAvis|u02_intHistoCommissionDecision|u02_txtHistoCommissionMotif|u02_booHistoCommissionFratrie
'           Ex: 03/01/2005|N|03/01/2005|03/03/2006|09/03/2006|03/10/2005|1|1|Projet insuffisamment élaboré|Non
        Case daoEXTRACTION_002
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT TOP 1 " _
                          & "DEMA.DEMA_DATE, DEMA.DEMA_CFRAT, DEMA.DEMA_DREF, " _
                          & "u02_HistoCommissions.u02_dtmHistoCommissionReunion, u02_HistoCommissions.u02_dtmHistoCommissionSignature, u02_HistoCommissions.u02_dtmHistoCommissionPrevu, u02_HistoCommissions.u02_intHistoCommissionAvis, u02_HistoCommissions.u02_intHistoCommissionDecision, u02_HistoCommissions.u02_txtHistoCommissionMotif, u02_HistoCommissions.u02_booHistoCommissionFratrie " _
                          & "FROM DEMA LEFT JOIN u02_HistoCommissions ON DEMA.DEMA_NUM = u02_HistoCommissions.DEMA_NUM " _
                          & "WHERE (DEMA.DEMA_NUM " & pstrArguments(0) & ") " _
                          & "ORDER BY u02_HistoCommissions.u02_dtmHistoCommissionReunion DESC"
    
'       frm_DemandeAdoptionFiche: Extraire les origines ethniques souhaitées pour l'enfant
'           ORET_CODE|ENDE_NUM
'           Ex: 2|3
        Case daoEXTRACTION_003
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT * " _
                          & "FROM ORSO " _
                          & "WHERE (ORSO.ENDE_NUM " & pstrArguments(0) & ")"

'       frm_DemandeAdoptionFiche: Extraire la demande
'           NATU_CODE|TYPE_CODE|DEMA_DATE|DEMA_CDDAM|DEMA_CFRAT|DEMA_DSORT|DEMA_MSORT|DEMA_REMA|DEMA_InclureReactualisation|TYPE_LIBELLE|NATU_LIBELLE
'           Ex: 6|6|24/07/1998|N|N|27/04/2004|ETRANGER 2004|1 garçon 5 ans Madagascar par les Enfants Avant To|Oui|Indifférent|Adoption plénière
        Case daoEXTRACTION_004
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "DEMA.NATU_CODE, DEMA.TYPE_CODE, DEMA.DEMA_DATE, DEMA.DEMA_CDDAM, DEMA.DEMA_CFRAT, DEMA.DEMA_DSORT, DEMA.DEMA_MSORT, DEMA.DEMA_REMA, DEMA.DEMA_InclureReactualisation, " _
                          & "TYPE.TYPE_LIBELLE, " _
                          & "NATU.NATU_LIBELLE " _
                          & "FROM DEMA, TYPE, NATU " _
                          & "WHERE (DEMA.DEMA_NUM " & pstrArguments(0) & " AND TYPE.TYPE_CODE = DEMA.TYPE_CODE AND NATU.NATU_CODE = DEMA.NATU_CODE)"
    
'       frm_RecalculEcheancier: Extraire la dernière/prochaine commission liée à la demande
'           DEMA_NUM|u02_dtmHistoCommissionReunion|u02_dtmHistoCommissionPrevu|u02_intHistoCommissionDecision
'           Ex: 9908540|22/09/2008|22/09/2008|1
        Case daoEXTRACTION_005
            plngTypeRecordset = dbOpenDynaset
            pstrChaîneSQL = "SELECT TOP 1 " _
                          & "u02_HistoCommissions.DEMA_NUM, u02_HistoCommissions.u02_dtmHistoCommissionReunion, u02_HistoCommissions.u02_dtmHistoCommissionPrevu, u02_HistoCommissions.u02_intHistoCommissionDecision " _
                          & "FROM u02_HistoCommissions " _
                          & "WHERE (u02_HistoCommissions.DEMA_NUM " & pstrArguments(0) & ") " _
                          & "ORDER BY u02_HistoCommissions.DEMA_NUM DESC, u02_HistoCommissions.u02_dtmHistoCommissionReunion DESC"

'       frm_RecalculEcheancier: Extraire les dates prévues pour la réception des rapports des intervenants
'           DEMA_NUM|TRSO_NUMAS|TRSO_NUMPSY|DEMA_DDEMRP|DEMA_DPRRP|DEMA_DDEMRA|DEMA_DPRRA|DEMA_DSORT|DEMA_DSUSP|DEMA_DREF
'           Ex: 9908540|215|184|29/11/2007|22/07/2008|29/11/2007|22/07/2008|||29/10/2007
        Case daoEXTRACTION_006
            plngTypeRecordset = dbOpenDynaset
            pstrChaîneSQL = "SELECT " _
                          & "DEMA.DEMA_NUM, DEMA.TRSO_NUMAS, DEMA.TRSO_NUMPSY, DEMA.DEMA_DDEMRP, DEMA.DEMA_DPRRP, DEMA.DEMA_DDEMRA, DEMA.DEMA_DPRRA, DEMA.DEMA_DSORT, DEMA.DEMA_DSUSP, DEMA.DEMA_DREF " _
                          & "FROM DEMA " _
                          & "WHERE (DEMA.DEMA_NUM " & pstrArguments(0) & ")"

'       frm_RecalculEcheancier: Extraire les intervenants
'           DEMA_NUM|u01_booHistoIntervenantPsychologue|u01_dtmHistoIntervenantDemande|u01_intHistoIntervenantNumero|TRSO_NUM|u01_dtmHistoIntervenantPrevu
'           Ex: 9907891|Oui|04/07/2002|2|184|06/02/2003
        Case daoEXTRACTION_007
            plngTypeRecordset = dbOpenDynaset
            pstrChaîneSQL = "SELECT " _
                          & "u01_HistoIntervenants.DEMA_NUM, u01_HistoIntervenants.u01_booHistoIntervenantPsychologue, u01_HistoIntervenants.u01_dtmHistoIntervenantDemande, u01_HistoIntervenants.u01_intHistoIntervenantNumero, u01_HistoIntervenants.TRSO_NUM, u01_HistoIntervenants.u01_dtmHistoIntervenantPrevu " _
                          & "FROM u01_HistoIntervenants " _
                          & "WHERE (u01_HistoIntervenants.DEMA_NUM " & pstrArguments(0) & " AND u01_HistoIntervenants.u01_booHistoIntervenantPsychologue = TRUE AND u01_HistoIntervenants.u01_dtmHistoIntervenantDemande " & pstrArguments(1) & " AND u01_HistoIntervenants.TRSO_NUM " & pstrArguments(2) & ") OR (u01_HistoIntervenants.DEMA_NUM " & pstrArguments(0) & " AND u01_HistoIntervenants.u01_booHistoIntervenantPsychologue = FALSE AND u01_HistoIntervenants.u01_dtmHistoIntervenantDemande " & pstrArguments(3) & " AND u01_HistoIntervenants.TRSO_NUM " & pstrArguments(4) & ") " _
                          & "ORDER BY u01_HistoIntervenants.u01_booHistoIntervenantPsychologue ASC, u01_HistoIntervenants.u01_intHistoIntervenantNumero DESC"
    
'       frm_RecalculEcheancier: Extraire le recours administratif
'           DEMA_NUM|u02_dtmHistoCommissionReunion|READ_DDREC|READ_DPREDEC|TRSO_NUMAS|READ_DDEMAS|READ_DPRRECAS|TRSO_NUMPSY|READ_DDEMPSY|READ_DPRRECPSY
'           Ex: 9908336|01/12/2006|21/01/2007|21/10/2007||||211|09/02/2007|09/09/2007
        Case daoEXTRACTION_008
            plngTypeRecordset = dbOpenDynaset
            pstrChaîneSQL = "SELECT " _
                          & "READ.DEMA_NUM, READ.u02_dtmHistoCommissionReunion, READ.READ_DDREC, READ.READ_DPREDEC, READ.TRSO_NUMAS, READ.READ_DDEMAS, READ.READ_DPRRECAS, READ.TRSO_NUMPSY, READ.READ_DDEMPSY, READ.READ_DPRRECPSY " _
                          & "FROM READ " _
                          & "WHERE (READ.DEMA_NUM " & pstrArguments(0) & " AND READ.u02_dtmHistoCommissionReunion " & pstrArguments(1) & ")"
    
'       frmDemandesÀRéactualiser: Extraire les demandes à réactualiser
'           DOSS_NUM|DEMA_NUM|INDI_NOM|INDI_PREN
'           Ex: 2008004|9908567|RUELLAN|Dominique
        Case daoEXTRACTION_009
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "DEMA.DOSS_NUM, DEMA.DEMA_NUM, " _
                          & "INDI.INDI_NOM, INDI.INDI_PREN " _
                          & "FROM (DOSS INNER JOIN DEMA ON DOSS.DOSS_NUM = DEMA.DOSS_NUM) LEFT JOIN INDI ON DOSS.DOSS_NUM = INDI.DOSS_NUM " _
                          & "WHERE (DEMA.DEMA_DSORT IS NULL AND DEMA.DEMA_InclureReactualisation = TRUE AND DEMA.DEMA_dtmAgrementDelivrance IS NOT NULL AND DEMA.u02_dtmHistoCommissionReunion IS NULL AND DateAdd('m', 20, DEMA_dtmAgrementDelivrance) < " & strSQLDateFormatée(Date) & " AND INDI.INDI_CDFR = '" & gstrTYPE_DEMANDEUR & "' AND NOT EXISTS (SELECT PLAC.DEMA_NUM FROM PLAC WHERE PLAC.DEMA_NUM = DEMA.DEMA_NUM) AND NOT EXISTS (SELECT RENO.DEMA_NUM FROM RENO WHERE RENO.DEMA_NUM = DEMA.DEMA_NUM AND RENO.RENO_DEMREA = '1' AND (RENO.RENO_DDEMMAAS IS NOT NULL OR RENO.RENO_DDEMMPS IS NOT NULL))) " & "ORDER BY DEMA.DOSS_NUM DESC" _
                              
'       frmEnvoiDossierTGI: Extraire la date d'envoi
'           PLAC_NUM|PLAC_dtmEnvoiTGI
'           Ex: 783|
        Case daoEXTRACTION_010
            plngTypeRecordset = dbOpenDynaset
            pstrChaîneSQL = "SELECT " _
                          & "PLAC.PLAC_NUM, PLAC.PLAC_dtmEnvoiTGI " _
                          & "FROM PLAC " _
                          & "WHERE (PLAC.PLAC_NUM " & pstrArguments(0) & ")"

'       frm_Individu: Extraire les demandeurs portant mêmes nom et prénom
'           DOSS_NUM
'           Ex: 1990022
        Case daoEXTRACTION_011
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "INDI.DOSS_NUM " _
                          & "FROM INDI " _
                          & "WHERE (INDI.INDI_CDFR = '" & gstrTYPE_DEMANDEUR & "' AND INDI.INDI_NOM " & pstrArguments(0) & " AND INDI.INDI_PREN " & pstrArguments(1) & " AND NOT INDI.INDI_NUM " & pstrArguments(2) & ")"
    
'       modDAO: Extraire les types de recours contentieux
'           TYRE_NUM|TYRE_NOM
'           Ex: 1|Tribunal Administratif
        Case daoEXTRACTION_012
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "TYRE.TYRE_NUM, TYRE.TYRE_NOM " _
                          & "FROM TYRE"

'       frm_DecisionCommission: Extraire les informations relatives à l'agrément si elles existent
'           DEMA_dtmAgrementDelivrance|DEMA_intAgrementProvenance|u02_dtmHistoCommissionReunion
'           Ex: 10/11/1981|-2|
        Case daoEXTRACTION_013
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "DEMA.DEMA_dtmAgrementDelivrance, DEMA.DEMA_intAgrementProvenance, DEMA.u02_dtmHistoCommissionReunion " _
                          & "FROM DEMA " _
                          & "WHERE (DEMA.DEMA_NUM " & pstrArguments(0) & " AND DEMA.DEMA_dtmAgrementDelivrance IS NOT NULL)"

'       frmHistoriqueCommissions: Extraire l'historique des commissions
'           u02_dtmHistoCommissionReunion|u02_dtmHistoCommissionSignature|u02_intHistoCommissionAvis|u02_intHistoCommissionDecision|u02_txtHistoCommissionMotif|u02_booHistoCommissionFratrie
'           Ex: 14/10/2005|21/10/2005|1|1|projet d'adoption considéré comme une mission|Non
        Case daoEXTRACTION_014
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "u02_HistoCommissions.u02_dtmHistoCommissionReunion, u02_HistoCommissions.u02_dtmHistoCommissionSignature, u02_HistoCommissions.u02_intHistoCommissionAvis, u02_HistoCommissions.u02_intHistoCommissionDecision, u02_HistoCommissions.u02_txtHistoCommissionMotif, u02_HistoCommissions.u02_booHistoCommissionFratrie " _
                          & "FROM u02_HistoCommissions " _
                          & "WHERE (u02_HistoCommissions.DEMA_NUM " & pstrArguments(0) & ") " _
                          & "ORDER BY u02_HistoCommissions.u02_dtmHistoCommissionReunion DESC"

'       frm_DecisionCommission: Extraire les enfants accordés pour la demande
'           ENDE_NUM
'           Ex: 5024
        Case daoEXTRACTION_015
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "COAC.ENDE_NUM " _
                          & "FROM COAC " _
                          & "WHERE (COAC.DEMA_NUM " & pstrArguments(0) & ") "
    
'       Frm_Recours: Extraire les refus et les retraits décidés par les commissions
'           u02_dtmHistoCommissionReunion|u02_intHistoCommissionDecision|u02_dtmHistoCommissionSignature
'           Ex: 01/01/2006|2|01/01/2006
        Case daoEXTRACTION_016
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "u02_HistoCommissions.u02_dtmHistoCommissionReunion, u02_HistoCommissions.u02_intHistoCommissionDecision, u02_HistoCommissions.u02_dtmHistoCommissionSignature " _
                          & "FROM u02_HistoCommissions " _
                          & "WHERE (u02_HistoCommissions.DEMA_NUM " & pstrArguments(0) & " AND (u02_HistoCommissions.u02_intHistoCommissionDecision = " & CStr(adpREFUS) & " OR u02_HistoCommissions.u02_intHistoCommissionDecision = " & CStr(adpRETRAIT) & ")) " _
                          & "ORDER BY u02_HistoCommissions.u02_dtmHistoCommissionReunion DESC"
    
'       modAdoptions: Extraire les numéros de dossiers
'           DOSS_NUM
'           Ex: 2008015
        Case daoEXTRACTION_017
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "DOSS.DOSS_NUM " _
                          & "FROM DOSS " _
                          & "ORDER BY DOSS.DOSS_NUM DESC"

'       modAdoptions: Extraire les numéros de demandes du dossier
'           DEMA_NUM
'           Ex: 9908581
        Case daoEXTRACTION_018
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "DEMA.DEMA_NUM " _
                          & "FROM DEMA " _
                          & "WHERE (DEMA.DOSS_NUM " & pstrArguments(0) & ")"

'       Frm_Recours: Extraire la date de dépôt de la demande
'           DEMA_DATE
'           Ex: 24/07/1998
        Case daoEXTRACTION_019
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "DEMA.DEMA_DATE " _
                          & "FROM DEMA " _
                          & "WHERE (DEMA.DEMA_NUM " & pstrArguments(0) & ")"
    
'       frm_DecisionCommission: Extraire les origines ethniques (souhaitées ou accordées)
'           ORET_CODE
'           Ex: 1
        Case daoEXTRACTION_020
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT DISTINCT " _
                          & "ORSO.ORET_CODE " _
                          & "FROM ORSO, " & pvarArgument(0) & " " _
                          & "WHERE (ORSO.ENDE_NUM = " & pvarArgument(0) & ".ENDE_NUM AND " & pvarArgument(0) & ".DEMA_NUM " & pstrArguments(1) & ")"
    
'       modAdoptions: Extraire les origines ethniques
'           ORET_CODE|ORET_LIBELLE
'           Ex: 1|Européenne
        Case daoEXTRACTION_021
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "ORET.ORET_CODE, ORET.ORET_LIBELLE " _
                          & "FROM ORET " _
                          & "ORDER BY ORET.ORET_CODE ASC"

'       frm_DemandeAdoptionEnfant: Extraire l'enfant (souhaité ou accordé)
'           ENDE_NUM|ENDE_SEXE|ENDE_HAND|ENDE_TRAGE
'           Ex: 5108|I|N|00-01
        Case daoEXTRACTION_022
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT * " _
                          & "FROM ENDE " _
                          & "WHERE (ENDE.ENDE_NUM " & pstrArguments(0) & ")"

'       frmRecoursContentieux: Extraire la décision de la commission
'           u02_intHistoCommissionDecision|u02_dtmHistoCommissionSignature
'           Ex: 2|01/01/2006
        Case daoEXTRACTION_023
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "u02_HistoCommissions.u02_intHistoCommissionDecision, u02_HistoCommissions.u02_dtmHistoCommissionSignature " _
                          & "FROM u02_HistoCommissions " _
                          & "WHERE (u02_HistoCommissions.DEMA_NUM " & pstrArguments(0) & " AND u02_HistoCommissions.u02_dtmHistoCommissionReunion " & pstrArguments(1) & ")"
    
'       modAdoptions: Extraire l'historique des intervenants
'           DEMA_NUM|u01_intHistoIntervenantCadre|u01_booHistoIntervenantPsychologue|u01_dtmHistoIntervenantDemande|u01_intHistoIntervenantNumero|u01_dtmHistoIntervenantPrevu|u01_dtmHistoIntervenantRecu|TRSO_NUM
'           Ex: 9908427|4|Oui|20/05/2008|1|20/12/2008||12
        Case daoEXTRACTION_024
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT * " _
                          & "FROM u01_HistoIntervenants " _
                          & "WHERE (u01_HistoIntervenants.DEMA_NUM " & pstrArguments(0) & ") " _
                          & "ORDER BY u01_HistoIntervenants.u01_booHistoIntervenantPsychologue ASC, u01_HistoIntervenants.u01_dtmHistoIntervenantDemande DESC, u01_HistoIntervenants.u01_intHistoIntervenantNumero DESC"
    
'       frm_DemandeAdoptionInstruction: Extraire la demande
'           DEMA_NUM|TRSO_NUMAS|TRSO_NUMPSY|DEMA_DATE|DEMA_DDEMAN|DEMA_DRECAN|DEMA_DDEMCJ|DEMA_DRECCJ|DEMA_DDEMCM|DEMA_DRECCM|DEMA_DDEMPS|DEMA_DRECPS|DEMA_DDEMNP|DEMA_DRECNP|DEMA_DDEMJR|DEMA_DRECJR|DEMA_DDEMRP|DEMA_DPRRP|DEMA_DRECRP|DEMA_DDEMRA|DEMA_DPRRA|DEMA_DRECRA|DEMA_DDEMFF|DEMA_DRECFF|DEMA_DPRDOS|DEMA_DREDOS|DEMA_DREF|TRSO_NUMPSY_Initial|PsychologueInitialNom|DEMA_DDEMRP_Initial|DEMA_DPRRP_Initial|DEMA_DRECRP_Initial|TRSO_NUMAS_Initial|AssistanteSocialeInitialeNom|DEMA_DDEMRA_Initial|DEMA_DPRRA_Initial|DEMA_DRECRA_Initial
'           Ex: 9908594|209|12|16/04/2008|16/04/2008|16/04/2008|16/04/2008|16/04/2008|16/04/2008|16/04/2008|||||16/04/2008|16/04/2008|29/04/2008|16/12/2008||29/04/2008|16/12/2008||16/04/2008|16/04/2008|16/05/2008|16/04/2008|16/04/2008||||||||||
        Case daoEXTRACTION_025
            plngTypeRecordset = dbOpenDynaset
            pstrChaîneSQL = "SELECT " _
                          & "DEMA.DEMA_NUM, DEMA.TRSO_NUMAS, DEMA.TRSO_NUMPSY, DEMA.DEMA_DATE, DEMA.DEMA_DDEMAN, DEMA.DEMA_DRECAN, DEMA.DEMA_DDEMCJ, DEMA.DEMA_DRECCJ, DEMA.DEMA_DDEMCM, DEMA.DEMA_DRECCM, DEMA.DEMA_DDEMPS, DEMA.DEMA_DRECPS, DEMA.DEMA_DDEMNP, DEMA.DEMA_DRECNP, DEMA.DEMA_DDEMJR, DEMA.DEMA_DRECJR, DEMA.DEMA_DDEMRP, DEMA.DEMA_DPRRP, DEMA.DEMA_DRECRP, DEMA.DEMA_DDEMRA, DEMA.DEMA_DPRRA, DEMA.DEMA_DRECRA, DEMA.DEMA_DDEMFF, DEMA.DEMA_DRECFF, DEMA.DEMA_DPRDOS, DEMA.DEMA_DREDOS, DEMA.DEMA_DREF, DEMA.TRSO_NUMPSY_Initial, DEMA.DEMA_DDEMRP_Initial, DEMA.DEMA_DPRRP_Initial, DEMA.DEMA_DRECRP_Initial, DEMA.TRSO_NUMAS_Initial, DEMA.DEMA_DDEMRA_Initial, DEMA.DEMA_DPRRA_Initial, DEMA.DEMA_DRECRA_Initial " _
                          & "FROM DEMA " _
                          & "WHERE (DEMA.DEMA_NUM " & pstrArguments(0) & ")"
    
'       Extraire un intervenant
'           TRSO_NOM
'           Ex: COLLIN
        Case daoEXTRACTION_026
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "TRSO.TRSO_NOM " _
                          & "FROM TRSO " _
                          & "WHERE (TRSO.TRSO_NUM " & pstrArguments(0) & ")"
                          

'       modAdoptions: Extraire l'historique des intervenants liés à un même cadre d'intervention
'           DEMA_NUM|u01_intHistoIntervenantCadre|u01_booHistoIntervenantPsychologue|u01_dtmHistoIntervenantDemande|u01_intHistoIntervenantNumero|u01_dtmHistoIntervenantPrevu|u01_dtmHistoIntervenantRecu|TRSO_NUM
'           Ex: 9908491|4|Oui|30/04/2008|1|08/12/2008||184
        Case daoEXTRACTION_027
            plngTypeRecordset = dbOpenDynaset
            pstrChaîneSQL = "SELECT * " _
                          & "FROM u01_HistoIntervenants " _
                          & "WHERE (u01_HistoIntervenants.DEMA_NUM " & pstrArguments(0) & " AND u01_HistoIntervenants.u01_intHistoIntervenantCadre " & pstrArguments(1) & " AND u01_HistoIntervenants.u01_intHistoIntervenantNumero " & pstrArguments(2) & ") " _
                          & "ORDER BY u01_HistoIntervenants.u01_booHistoIntervenantPsychologue, u01_HistoIntervenants.u01_dtmHistoIntervenantDemande DESC"
    
'       frm_DemandeAdoptionInstruction: Extraire la demande
'           DEMA_NUM|DEMA_DATE|DEMA_DSUSP|DEMA_MSUSP|DEMA_DREPR
'           Ex: 8804501|14/12/1988|09/02/1990|NAISSANCE ATTENDUE + CHGT PROF MR L|10/12/1990
        Case daoEXTRACTION_028
            plngTypeRecordset = dbOpenDynaset
            pstrChaîneSQL = "SELECT " _
                          & "DEMA.DEMA_NUM, DEMA.DEMA_DATE, DEMA.DEMA_DSUSP, DEMA.DEMA_MSUSP, DEMA_DREPR " _
                          & "FROM DEMA " _
                          & "WHERE (DEMA.DEMA_NUM " & pstrArguments(0) & ")"

'       modAdoptions: Extraire le numéro de dossier de la demande
'           DOSS_NUM
'           Ex: 2008040
        Case daoEXTRACTION_029
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "DEMA.DOSS_NUM " _
                          & "FROM DEMA " _
                          & "WHERE (DEMA.DEMA_NUM " & pstrArguments(0) & ")"

'       frmRecoursContentieux: Extraire la date de réunion de la première commission à prendre une décision
'           DatePremiereCommission
'           Ex: 03/06/2008
        Case daoEXTRACTION_031
            plngTypeRecordset = dbOpenSnapshot
            pstrChaîneSQL = "SELECT " _
                          & "MIN(u02_HistoCommissions.u02_dtmHistoCommissionReunion) AS DatePremiereCommission " _
                          & "FROM u02_HistoCommissions " _
                          & "WHERE (u02_HistoCommissions.DEMA_NUM " & pstrArguments(0) & " AND NOT u02_HistoCommissions.u02_intHistoCommissionDecision = " & CStr(AdoptionsDélibérations.adpINDÉTERMINÉ) & ")"

'       frm_PlacementAjoutTS: Compter les demandes de rapports se rapportant à autre chose qu'à une surveillance de placement
'           Nombre
'           Ex: 1
        Case daoEXTRACTION_032
            plngTypeRecordset = dbOpenForwardOnly
            pstrChaîneSQL = "SELECT " _
                          & "COUNT(*) AS Nombre " _
                          & "FROM u01_HistoIntervenants " _
                          & "WHERE (u01_HistoIntervenants.DEMA_NUM " & pstrArguments(0) & " AND NOT u01_HistoIntervenants.u01_intHistoIntervenantCadre " & pstrArguments(1) & " AND u01_HistoIntervenants.u01_booHistoIntervenantPsychologue " & pstrArguments(2) & " AND u01_HistoIntervenants.u01_dtmHistoIntervenantDemande " & pstrArguments(3) & ")"

'       Requête pour test de verrouillage
        Case 50000
            plngTypeRecordset = dbOpenDynaset
            pstrChaîneSQL = "SELECT * FROM RENO WHERE RENO.CONF_ANNEE = 1 AND RENO.DEMA_NUM = 9908507"
    
    End Select

    pdatHeureRéférence = Now
    
    Do
    
'       Si les données ont été verrouillées par une transaction
        If pbooVerrouillageActivé Then
'           Avertir l'utilisateur du verrouillage au cas où le délai maximal avant l'affichage d'un message serait atteint
            GérerVerrouillage pbooVerrouillageActivé, pdatHeureRéférence
'           Annuler l'opération si l'utilisateur le demande
            If mlngDernièreErreurDAO = daoENREGISTREMENT_VERROUILLÉ Then
'               Si un message supplémentaire doit être affiché
                If Not pstrListesMessages(pkintVerrouillage) = "" Then
                    pstrMessages() = Split(pstrListesMessages(pkintVerrouillage), gstrSÉPARATEUR_2)
'                   Si aucun pictogramme n'est défini
                    If UBound(pstrMessages) = 0 Then
                        ReDim Preserve pstrMessages(1)
                        pstrMessages(1) = vbExclamation
                    End If
                    Err.Raise mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION
                End If
                If pbooGérerSablier Then Screen.MousePointer = vbDefault
                Exit Function
            End If
        End If
        
        mlngDernièreErreurDAO = daoAUCUNE
        
'       Extraire le jeu d'enregistrements
#If TablesEnLocalOuAttachées Then
        Set rstSQLExtraction = mconnODBCAdoptions.OpenRecordset(pstrChaîneSQL, plngTypeRecordset, 0, dbOptimisticBatch)
#Else
        If plngTypeRecordset = dbOpenDynaset Then
            Set rstSQLExtraction = mconnODBCAdoptions.OpenRecordset(pstrChaîneSQL, plngTypeRecordset, 0, dbPessimistic)
        Else
            Set rstSQLExtraction = mconnODBCAdoptions.OpenRecordset(pstrChaîneSQL, plngTypeRecordset)
        End If
#End If
        DoEvents

'   Boucler tant qu'une transaction verrouille les données
    Loop While pbooVerrouillageActivé

    mstrChaîneSQL = pstrChaîneSQL: mlngTypeRecordset = plngTypeRecordset
    
    RemplirRecordsetEnTotalité rstSQLExtraction
    
'   Si le jeu ne contient pas d'enregistrement
    If mlngDernièreErreurDAO = daoAUCUN_ENREGISTREMENT Then
'       En informer l'utilisateur si cela est demandé
        If Not pstrListesMessages(pkintRecordsetVide) = "" Then Err.Raise mesRECORDSET_VIDE
        If pbooGérerSablier Then Screen.MousePointer = vbDefault
        Exit Function
    End If

    If pbooGérerSablier Then Screen.MousePointer = vbDefault
    
    Exit Function

Err_rstSQLExtraction:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                                     pintTraitement = mesINTERROMPRE
        Case daoERREUR_03146:                                   If mdbeAdoptions.Errors(0).Number = daoERREUR_03261 Then pbooVerrouillageActivé = True: Resume Next
                                                                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
        Case mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrMessages(0), pstrMessages(1))
        Case mesRECORDSET_VIDE:                                 pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrListesMessages(pkintRecordsetVide), pstrChaîneSQL)
        Case Else:                                              pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Function


'Ajouter un enregistrement
'   Les valeurs prises par mlngDernièreErreurDAO sont:  (*) daoENREGISTREMENT_VERROUILLÉ:       Les données sont verrouillées par une transaction
'                                                       (#) daoDOUBLON:                         L'ajout a été refusé pour éviter la création d'un doublon
'                                                           daoENREGISTREMENT_CONNEXE_MANQUANT: L'ajout a été refusé du fait d'un enregistrement connexe manquant
'                                                       (*) daoAUCUNE:                          L'enregistrement a été ajouté
'       (*) la valeur est retournée et peut donc être interceptée par la procédure appelante
'       (#) la valeur n'est retournée que si un message est défini (message si doublon); si le message correspond à gstrDOUBLON_SANS_MESSAGE, la valeur est retourné sans aucun message
'   plngNuméroRequête:  numéro de la requête
'   pbooGérerSablier:   booléen précisant s'il faut gérer le sablier
'   pstrListeMessages:  liste de messages déterminant les messages supplémentaires à afficher en cas d'erreur au format '<message si verrouillage[¤pictogramme]>[|<message si doublon[¤pictogramme]>]'
'                       pour chaque message, 3 formats possibles:   "":                         aucun message n'est affiché
'                                                                   "<message>":                le message est affiché avec le pictogramme vbExclamation
'                                                                   "<message>¤<pictogramme>":  le message est affiché avec le pictogramme désigné
'   pvarArgument():     tableau d'arguments

Public Sub SQLAjouter(plngNuméroRequête As DAOAjouts, pbooGérerSablier As Boolean, ByVal pstrListeMessages As String, ParamArray pvarArgument() As Variant)
Const pkstrNomProcédure As String = "SQLAjouter"
Const pkintVerrouillage As Integer = 0, pkintDoublon As Integer = 1
Dim pdatHeureRéférence As Date
Dim pstrChaîneSQL As String, pstrArguments() As String, pstrListesMessages() As String, pstrMessages() As String, pstrNomTable As String
Dim pintIndiceArgument As Integer, pintPositionApostrophe() As Integer, pintNuméroApostrophe() As Integer, pintIndiceTableau As Integer, pintPositionDépartRecherche As Integer
Dim pbooVerrouillageActivé As Boolean
On Error GoTo Err_SQLAjouter

    If pbooGérerSablier Then Screen.MousePointer = vbHourglass
    
    pstrListesMessages() = Split(pstrListeMessages, gstrSÉPARATEUR_1)
    While Not UBound(pstrListesMessages) = pkintDoublon
        pstrListeMessages = pstrListeMessages & gstrSÉPARATEUR_1
        pstrListesMessages() = Split(pstrListeMessages, gstrSÉPARATEUR_1)
    Wend
    
    If Not UBound(pvarArgument) = comINDÉFINI Then ReDim pstrArguments(LBound(pvarArgument) To UBound(pvarArgument))
    
'   Pour chaque argument passé en référence
    For pintIndiceArgument = LBound(pvarArgument) To UBound(pvarArgument)
        Select Case VarType(pvarArgument(pintIndiceArgument))
            Case vbBoolean
                pstrArguments(pintIndiceArgument) = IIf(pvarArgument(pintIndiceArgument), "TRUE", "FALSE")
            Case vbDate
                pstrArguments(pintIndiceArgument) = strSQLDateFormatée(pvarArgument(pintIndiceArgument))
            Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                pstrArguments(pintIndiceArgument) = strSQLNombreFormaté(pvarArgument(pintIndiceArgument))
            Case vbNull
                pstrArguments(pintIndiceArgument) = "NULL"
            Case vbString
                pstrArguments(pintIndiceArgument) = "'" & strSQLChaîneFormatée(pvarArgument(pintIndiceArgument)) & "'"
        End Select
    Next pintIndiceArgument

    Select Case plngNuméroRequête

'       Création d'un recours administratif
        Case daoADD_RECOURS_ADMINISTRATIF
            pstrChaîneSQL = "INSERT INTO READ (TRSO_NUMAS, TRSO_NUMPSY, READ_DDREC, READ_DCOM, READ_DPREDEC, READ_DDECI, READ_DECIC, READ_DECIF, READ_DDEMAS, READ_DPRRECAS, READ_RECAS, READ_DDEMPSY, READ_DPRRECPSY, READ_RECPSY, READ_REFUS, DEMA_NUM, u02_dtmHistoCommissionReunion) " _
                          & "VALUES (" & pstrArguments(0) & ", " & pstrArguments(1) & ", " & pstrArguments(2) & ", " & pstrArguments(3) & ", " & pstrArguments(4) & ", " & pstrArguments(5) & ", " & pstrArguments(6) & ", " & pstrArguments(7) & ", " & pstrArguments(8) & ", " & pstrArguments(9) & ", " & pstrArguments(10) & ", " & pstrArguments(11) & ", " & pstrArguments(12) & ", " & pstrArguments(13) & ", " & pstrArguments(14) & ", " & pstrArguments(15) & ", " & pstrArguments(16) & ")"
    
'       Création d'un enfant
        Case daoADD_ENFANT
            pstrChaîneSQL = "INSERT INTO ENDE (ENDE_NUM, ENDE_SEXE, ENDE_HAND, ENDE_TRAGE) " _
                          & "VALUES (" & pstrArguments(0) & ", " & pstrArguments(1) & ", " & pstrArguments(2) & ", " & pstrArguments(3) & ")"
    
'       Création d'un enfant demandé ou accordé en commission
        Case daoADD_ENFANT_DEMANDÉ_OU_ACCORDÉ
            pstrChaîneSQL = "INSERT INTO " & pvarArgument(0) & " (DEMA_NUM, ENDE_NUM) " _
                          & "VALUES (" & pstrArguments(1) & ", " & pstrArguments(2) & ")"
    
'       Création d'un historique pour l'intervenant
        Case daoADD_HISTORIQUE_INTERVENANT
            pstrChaîneSQL = "INSERT INTO u01_HistoIntervenants (DEMA_NUM, u01_intHistoIntervenantCadre, u01_booHistoIntervenantPsychologue, u01_dtmHistoIntervenantDemande, u01_intHistoIntervenantNumero, u01_dtmHistoIntervenantPrevu, u01_dtmHistoIntervenantRecu, TRSO_NUM) " _
                          & "VALUES (" & pstrArguments(0) & ", " & pstrArguments(1) & ", " & pstrArguments(2) & ", " & pstrArguments(3) & ", " & pstrArguments(4) & ", " & pstrArguments(5) & ", " & pstrArguments(6) & ", " & pstrArguments(7) & ")"
    
'       frm_DecisionCommission: Création d'un historique pour la commission
        Case daoAJOUT_001
            pstrChaîneSQL = "INSERT INTO u02_HistoCommissions (DEMA_NUM, u02_dtmHistoCommissionReunion, u02_dtmHistoCommissionSignature, u02_dtmHistoCommissionPrevu, u02_intHistoCommissionAvis, u02_intHistoCommissionDecision, u02_txtHistoCommissionMotif, u02_booHistoCommissionFratrie) " _
                          & "VALUES (" & pstrArguments(0) & ", " & pstrArguments(1) & ", " & pstrArguments(2) & ", " & pstrArguments(3) & ", " & pstrArguments(4) & ", " & pstrArguments(5) & ", " & pstrArguments(6) & ", " & pstrArguments(7) & ")"
    
'       frm_DecisionCommission: Création d'un recours contentieux
        Case daoAJOUT_002
            pstrChaîneSQL = "INSERT INTO RECO (TYRE_NUM, RECO_DDREC, RECO_DPREDEC, RECO_DDECI, RECO_DECI, RECO_REFUS, RECO_dtmAgrementDelivrance, DEMA_NUM, u02_dtmHistoCommissionReunion) " _
                          & "VALUES (" & pstrArguments(0) & ", " & pstrArguments(1) & ", " & pstrArguments(2) & ", " & pstrArguments(3) & ", " & pstrArguments(4) & ", " & pstrArguments(5) & ", " & pstrArguments(6) & ", " & pstrArguments(7) & ", " & pstrArguments(8) & ")"
    
'       frm_DecisionCommission: Création de l'origine ethnique d'un enfant
        Case daoAJOUT_003
            pstrChaîneSQL = "INSERT INTO ORSO (ORET_CODE, ENDE_NUM) " _
                          & "VALUES (" & pstrArguments(0) & ", " & pstrArguments(1) & ")"
    
'       Requête pour test de verrouillage
        Case 50000
            pstrChaîneSQL = "INSERT INTO TYRE (TYRE_NUM, TYRE_NOM) " _
                          & "VALUES (50000, 'Verrou')"
    
    End Select

    pdatHeureRéférence = Now
    
    Do
    
'       Si les données ont été verrouillées par une transaction
        If pbooVerrouillageActivé Then
'           Avertir l'utilisateur du verrouillage au cas où le délai maximal avant l'affichage d'un message serait atteint
            GérerVerrouillage pbooVerrouillageActivé, pdatHeureRéférence
'           Annuler l'opération si l'utilisateur le demande
            If mlngDernièreErreurDAO = daoENREGISTREMENT_VERROUILLÉ Then
'               Si un message supplémentaire doit être affiché
                If Not pstrListesMessages(pkintVerrouillage) = "" Then
                    pstrMessages() = Split(pstrListesMessages(pkintVerrouillage), gstrSÉPARATEUR_2)
'                   Si aucun pictogramme n'est défini
                    If UBound(pstrMessages) = 0 Then
                        ReDim Preserve pstrMessages(1)
                        pstrMessages(1) = vbExclamation
                    End If
                    Err.Raise mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION
                End If
                If pbooGérerSablier Then Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        
        mlngDernièreErreurDAO = daoAUCUNE
        
'       Exécuter l'ajout
        mconnODBCAdoptions.Execute pstrChaîneSQL, dbExecDirect
        DoEvents
        
'   Boucler tant qu'une transaction verrouille les données
    Loop While pbooVerrouillageActivé
    
'   Si l'ajout a généré une erreur
    If Not mlngDernièreErreurDAO = daoAUCUNE Then
    
        Select Case mlngDernièreErreurDAO
        
'           L'ajout a été refusé pour éviter la création d'un doublon
            Case daoDOUBLON:
'               Si l'erreur n'est pas gérée
                If pstrListesMessages(pkintDoublon) = "" Then
                    Err.Raise mesANNULATION_SUR_DOUBLON
'               Sinon
                Else
'                   Si l'erreur est gérée après l'affichage d'un message
                    If Not pstrListesMessages(pkintDoublon) = gstrDOUBLON_SANS_MESSAGE Then
                        pstrMessages() = Split(pstrListesMessages(pkintDoublon), gstrSÉPARATEUR_2)
'                       Si aucun pictogramme n'est défini
                        If UBound(pstrMessages) = 0 Then
                            ReDim Preserve pstrMessages(1)
                            pstrMessages(1) = vbExclamation
                        End If
                        Err.Raise mesMESSAGE_AVEC_ICÔNE
                    End If
                End If
                
'           L'ajout a été refusé du fait d'un enregistrement connexe manquant
            Case daoENREGISTREMENT_CONNEXE_MANQUANT:
                ReDim pintPositionApostrophe(1 To ERREUR_3201_NOMBRE_APOSTROPHES) As Integer: ReDim pintNuméroApostrophe(0) As Integer: pintNuméroApostrophe(0) = CInt(ERREUR_3201_NUMÉRO_APOSTROPHE_AVANT_NOM_TABLE)
'               Extraire les positions des apostrophes dans le message
                For pintIndiceTableau = 1 To UBound(pintPositionApostrophe)
                    If pintIndiceTableau = 1 Then pintPositionDépartRecherche = 1 Else pintPositionDépartRecherche = pintPositionApostrophe(pintIndiceTableau - 1) + 1
                    pintPositionApostrophe(pintIndiceTableau) = InStr(pintPositionDépartRecherche, mdbeAdoptions.Errors(0).Description, "'")
                Next pintIndiceTableau
'               Extraire du message d'erreur DAO le nom de la table connexe
                pstrNomTable = Mid(mdbeAdoptions.Errors(0).Description, pintPositionApostrophe(pintNuméroApostrophe(0)) + 1, pintPositionApostrophe(pintNuméroApostrophe(0) + 1) - pintPositionApostrophe(pintNuméroApostrophe(0)) - 1)
                Err.Raise mesANNULATION_SUR_CONNEXE_MANQUANT
                
        End Select
        
        If pbooGérerSablier Then Screen.MousePointer = vbDefault
        Exit Sub
        
    End If

    If pbooGérerSablier Then Screen.MousePointer = vbDefault
        
    Exit Sub

Err_SQLAjouter:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                                     pintTraitement = mesINTERROMPRE
        Case daoERREUR_03146:                                   If mdbeAdoptions.Errors(0).Number = daoERREUR_03022 Then mlngDernièreErreurDAO = daoDOUBLON: Resume Next
                                                                If mdbeAdoptions.Errors(0).Number = daoERREUR_03201 Then mlngDernièreErreurDAO = daoENREGISTREMENT_CONNEXE_MANQUANT: Resume Next
                                                                If mdbeAdoptions.Errors(0).Number = daoERREUR_01304 Or mdbeAdoptions.Errors(0).Number = daoERREUR_03261 Then pbooVerrouillageActivé = True: Resume Next
                                                                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
        Case mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrMessages(0), pstrMessages(1))
        Case mesANNULATION_SUR_DOUBLON:                         pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, "L'ajout", "", pstrChaîneSQL)
        Case mesANNULATION_SUR_CONNEXE_MANQUANT:                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, "L'ajout", "", pstrNomTable, pstrChaîneSQL)
        Case mesMESSAGE_AVEC_ICÔNE:                             pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrMessages(0), pstrMessages(1))
        Case Else:                                              pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub


'Mettre à jour un enregistrement
'   Les valeurs prises par mlngDernièreErreurDAO sont:  (*) daoENREGISTREMENT_VERROUILLÉ:       Les données sont verrouillées par une transaction
'                                                       (#) daoDOUBLON:                         La mise à jour a été refusée pour éviter la création d'un doublon
'                                                           daoENREGISTREMENT_CONNEXE_EXISTANT: La mise à jour a été refusée du fait d'un enregistrement connexe existant
'                                                           daoENREGISTREMENT_CONNEXE_MANQUANT: La mise à jour a été refusée du fait d'un enregistrement connexe manquant
'                                                       (*) daoAUCUNE:                          L'enregistrement a été mis à jour
'       (*) la valeur est retournée et peut donc être interceptée par la procédure appelante
'       (#) la valeur n'est retournée que si un message est défini (message si doublon)
'   plngNuméroRequête:  numéro de la requête
'   pbooGérerSablier:   booléen précisant s'il faut gérer le sablier
'   pstrListeMessages:  liste de messages déterminant les messages supplémentaires à afficher en cas d'erreur au format '<message si verrouillage[¤pictogramme]>[|<message si doublon[¤pictogramme]>]'
'                       pour chaque message, 3 formats possibles:   "":                         aucun message n'est affiché
'                                                                   "<message>":                le message est affiché avec le pictogramme vbExclamation
'                                                                   "<message>¤<pictogramme>":  le message est affiché avec le pictogramme désigné
'   pvarArgument():     tableau d'arguments

Public Sub SQLMettreÀJour(plngNuméroRequête As DAOMisesÀJour, pbooGérerSablier As Boolean, ByVal pstrListeMessages As String, ParamArray pvarArgument() As Variant)
Const pkstrNomProcédure As String = "SQLMettreÀJour"
Const pkintVerrouillage As Integer = 0, pkintDoublon As Integer = 1
Dim pdatHeureRéférence As Date
Dim pstrChaîneSQL As String, pstrArguments() As String, pstrListesMessages() As String, pstrMessages() As String, pstrNomTable As String
Dim pintIndiceArgument As Integer, pintPositionApostrophe() As Integer, pintNuméroApostrophe() As Integer, pintIndiceTableau As Integer, pintPositionDépartRecherche As Integer
Dim pbooVerrouillageActivé As Boolean
On Error GoTo Err_SQLMettreÀJour

    If pbooGérerSablier Then Screen.MousePointer = vbHourglass
    
    pstrListesMessages() = Split(pstrListeMessages, gstrSÉPARATEUR_1)
    While Not UBound(pstrListesMessages) = pkintDoublon
        pstrListeMessages = pstrListeMessages & gstrSÉPARATEUR_1
        pstrListesMessages() = Split(pstrListeMessages, gstrSÉPARATEUR_1)
    Wend
    
    If Not UBound(pvarArgument) = comINDÉFINI Then ReDim pstrArguments(LBound(pvarArgument) To UBound(pvarArgument))
    
'   Pour chaque argument passé en référence
    For pintIndiceArgument = LBound(pvarArgument) To UBound(pvarArgument)
        Select Case VarType(pvarArgument(pintIndiceArgument))
            Case vbBoolean
                pstrArguments(pintIndiceArgument) = "= " & IIf(pvarArgument(pintIndiceArgument), "TRUE", "FALSE")
            Case vbDate
                pstrArguments(pintIndiceArgument) = "= " & strSQLDateFormatée(pvarArgument(pintIndiceArgument))
            Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                pstrArguments(pintIndiceArgument) = "= " & strSQLNombreFormaté(pvarArgument(pintIndiceArgument))
            Case vbNull
                pstrArguments(pintIndiceArgument) = "= NULL"
            Case vbString
                pstrArguments(pintIndiceArgument) = "= '" & strSQLChaîneFormatée(pvarArgument(pintIndiceArgument)) & "'"
        End Select
    Next pintIndiceArgument

    Select Case plngNuméroRequête

'       Mise à jour de l'historique pour l'intervenant
        Case daoMAJ_HISTORIQUE_INTERVENANT
            pstrChaîneSQL = "UPDATE u01_HistoIntervenants " _
                          & "SET u01_HistoIntervenants.u01_dtmHistoIntervenantDemande " & pstrArguments(4) & ", u01_HistoIntervenants.u01_dtmHistoIntervenantPrevu " & pstrArguments(5) & ", u01_HistoIntervenants.u01_dtmHistoIntervenantRecu " & pstrArguments(6) & ", u01_HistoIntervenants.TRSO_NUM " & pstrArguments(7) & " " _
                          & "WHERE (u01_HistoIntervenants.DEMA_NUM " & Replace(pstrArguments(0), "= NULL", "IS NULL") & " AND u01_HistoIntervenants.u01_booHistoIntervenantPsychologue " & Replace(pstrArguments(1), "= NULL", "IS NULL") & " AND u01_HistoIntervenants.u01_dtmHistoIntervenantDemande " & Replace(pstrArguments(2), "= NULL", "IS NULL") & " AND u01_HistoIntervenants.u01_intHistoIntervenantNumero " & Replace(pstrArguments(3), "= NULL", "IS NULL") & ")"

'       frm_DecisionCommission: Mise à jour d'un historique pour la commission
        Case daoMISE_À_JOUR_001
            pstrChaîneSQL = "UPDATE u02_HistoCommissions " _
                          & "SET u02_HistoCommissions.u02_dtmHistoCommissionReunion " & pstrArguments(2) & ", u02_HistoCommissions.u02_dtmHistoCommissionSignature " & pstrArguments(3) & ", u02_HistoCommissions.u02_dtmHistoCommissionPrevu " & pstrArguments(4) & ", u02_HistoCommissions.u02_intHistoCommissionAvis " & pstrArguments(5) & ", u02_HistoCommissions.u02_intHistoCommissionDecision " & pstrArguments(6) & ", u02_HistoCommissions.u02_txtHistoCommissionMotif " & pstrArguments(7) & ", u02_HistoCommissions.u02_booHistoCommissionFratrie " & pstrArguments(8) & " " _
                          & "WHERE (u02_HistoCommissions.DEMA_NUM " & Replace(pstrArguments(0), "= NULL", "IS NULL") & " AND u02_HistoCommissions.u02_dtmHistoCommissionReunion " & Replace(pstrArguments(1), "= NULL", "IS NULL") & ")"

'       modAdoptions: Mise à jour des données relatives à l'agrément
        Case daoMISE_À_JOUR_002
            pstrChaîneSQL = "UPDATE DEMA " _
                          & "SET DEMA.DEMA_dtmAgrementDelivrance " & pstrArguments(1) & ", DEMA.DEMA_intAgrementProvenance " & pstrArguments(2) & ", DEMA.u02_dtmHistoCommissionReunion " & pstrArguments(3) & " " _
                          & "WHERE (DEMA.DEMA_NUM " & Replace(pstrArguments(0), "= NULL", "IS NULL") & ")"

'       frm_DemandeAdoptionEnfant: Mise à jour de l'enfant (demandé ou accordé)
        Case daoMISE_À_JOUR_003
            pstrChaîneSQL = "UPDATE ENDE " _
                          & "SET ENDE.ENDE_SEXE " & pstrArguments(1) & ", ENDE.ENDE_HAND " & pstrArguments(2) & ", ENDE.ENDE_TRAGE " & pstrArguments(3) & " " _
                          & "WHERE (ENDE.ENDE_NUM " & Replace(pstrArguments(0), "= NULL", "IS NULL") & ")"

'       modAdoptions: Mise à jour des données relatives aux intervenants présents sur une demande
        Case daoMISE_À_JOUR_004
            pstrChaîneSQL = "UPDATE DEMA " _
                          & "SET DEMA.TRSO_NUMAS " & pstrArguments(1) & ", DEMA.TRSO_NUMPSY " & pstrArguments(2) & ", DEMA.DEMA_DDEMRP " & pstrArguments(3) & ", DEMA.DEMA_DPRRP " & pstrArguments(4) & ", DEMA.DEMA_DRECRP " & pstrArguments(5) & ", DEMA.DEMA_DDEMRA " & pstrArguments(6) & ", DEMA.DEMA_DPRRA " & pstrArguments(7) & ", DEMA.DEMA_DRECRA " & pstrArguments(8) & ", DEMA.TRSO_NUMPSY_Initial " & pstrArguments(9) & ", DEMA.DEMA_DDEMRP_Initial " & pstrArguments(10) & ", DEMA.DEMA_DPRRP_Initial " & pstrArguments(11) & ", DEMA.DEMA_DRECRP_Initial " & pstrArguments(12) & ", DEMA.TRSO_NUMAS_Initial " & pstrArguments(13) & ", DEMA.DEMA_DDEMRA_Initial " & pstrArguments(14) & ", DEMA.DEMA_DPRRA_Initial " & pstrArguments(15) & ", DEMA.DEMA_DRECRA_Initial " & pstrArguments(16) & " " _
                          & "WHERE (DEMA.DEMA_NUM " & Replace(pstrArguments(0), "= NULL", "IS NULL") & ")"
                          
    End Select

    pdatHeureRéférence = Now
    
    Do
    
'       Si les données ont été verrouillées par une transaction
        If pbooVerrouillageActivé Then
'           Avertir l'utilisateur du verrouillage au cas où le délai maximal avant l'affichage d'un message serait atteint
            GérerVerrouillage pbooVerrouillageActivé, pdatHeureRéférence
'           Annuler l'opération si l'utilisateur le demande
            If mlngDernièreErreurDAO = daoENREGISTREMENT_VERROUILLÉ Then
'               Si un message supplémentaire doit être affiché
                If Not pstrListesMessages(pkintVerrouillage) = "" Then
                    pstrMessages() = Split(pstrListesMessages(pkintVerrouillage), gstrSÉPARATEUR_2)
'                   Si aucun pictogramme n'est défini
                    If UBound(pstrMessages) = 0 Then
                        ReDim Preserve pstrMessages(1)
                        pstrMessages(1) = vbExclamation
                    End If
                    Err.Raise mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION
                End If
                If pbooGérerSablier Then Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        
        mlngDernièreErreurDAO = daoAUCUNE
        
'       Exécuter la mise à jour
        mconnODBCAdoptions.Execute pstrChaîneSQL, dbExecDirect
        DoEvents
        
'   Boucler tant qu'une transaction verrouille les données
    Loop While pbooVerrouillageActivé
    
'   Si l'ajout a généré une erreur
    If Not mlngDernièreErreurDAO = daoAUCUNE Then
    
        Select Case mlngDernièreErreurDAO
        
'           La mise à jour a été refusée pour éviter la création d'un doublon
            Case daoDOUBLON:
'               Si l'erreur n'est pas gérée
                If pstrListesMessages(pkintDoublon) = "" Then
                    Err.Raise mesANNULATION_SUR_DOUBLON
'               Sinon
                Else
                    pstrMessages() = Split(pstrListesMessages(pkintDoublon), gstrSÉPARATEUR_2)
'                   Si aucun pictogramme n'est défini
                    If UBound(pstrMessages) = 0 Then
                        ReDim Preserve pstrMessages(1)
                        pstrMessages(1) = vbExclamation
                    End If
                    Err.Raise mesMESSAGE_AVEC_ICÔNE
                End If

'           La mise à jour a été refusée du fait d'un enregistrement connexe manquant
            Case daoENREGISTREMENT_CONNEXE_MANQUANT:
                ReDim pintPositionApostrophe(1 To ERREUR_3201_NOMBRE_APOSTROPHES) As Integer: ReDim pintNuméroApostrophe(0) As Integer: pintNuméroApostrophe(0) = CInt(ERREUR_3201_NUMÉRO_APOSTROPHE_AVANT_NOM_TABLE)
'               Extraire les positions des apostrophes dans le message
                For pintIndiceTableau = 1 To UBound(pintPositionApostrophe)
                    If pintIndiceTableau = 1 Then pintPositionDépartRecherche = 1 Else pintPositionDépartRecherche = pintPositionApostrophe(pintIndiceTableau - 1) + 1
                    pintPositionApostrophe(pintIndiceTableau) = InStr(pintPositionDépartRecherche, mdbeAdoptions.Errors(0).Description, "'")
                Next pintIndiceTableau
'               Extraire du message d'erreur DAO le nom de la table connexe
                pstrNomTable = Mid(mdbeAdoptions.Errors(0).Description, pintPositionApostrophe(pintNuméroApostrophe(0)) + 1, pintPositionApostrophe(pintNuméroApostrophe(0) + 1) - pintPositionApostrophe(pintNuméroApostrophe(0)) - 1)
                Err.Raise mesANNULATION_SUR_CONNEXE_MANQUANT
                
'           La mise à jour a été refusée du fait d'un enregistrement connexe existant
            Case daoENREGISTREMENT_CONNEXE_EXISTANT:
                ReDim pintPositionApostrophe(1 To ERREUR_3200_NOMBRE_APOSTROPHES) As Integer: ReDim pintNuméroApostrophe(0) As Integer: pintNuméroApostrophe(0) = CInt(ERREUR_3200_NUMÉRO_APOSTROPHE_AVANT_NOM_TABLE)
'               Extraire les positions des apostrophes dans le message
                For pintIndiceTableau = 1 To UBound(pintPositionApostrophe)
                    If pintIndiceTableau = 1 Then pintPositionDépartRecherche = 1 Else pintPositionDépartRecherche = pintPositionApostrophe(pintIndiceTableau - 1) + 1
                    pintPositionApostrophe(pintIndiceTableau) = InStr(pintPositionDépartRecherche, mdbeAdoptions.Errors(0).Description, "'")
                Next pintIndiceTableau
'               Extraire du message d'erreur DAO le nom de la table connexe
                pstrNomTable = Mid(mdbeAdoptions.Errors(0).Description, pintPositionApostrophe(pintNuméroApostrophe(0)) + 1, pintPositionApostrophe(pintNuméroApostrophe(0) + 1) - pintPositionApostrophe(pintNuméroApostrophe(0)) - 1)
                Err.Raise mesANNULATION_SUR_CONNEXE_EXISTANT
                        
        End Select
        
        If pbooGérerSablier Then Screen.MousePointer = vbDefault
        Exit Sub
        
    End If

    If pbooGérerSablier Then Screen.MousePointer = vbDefault
        
    Exit Sub

Err_SQLMettreÀJour:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                                     pintTraitement = mesINTERROMPRE
        Case daoERREUR_03146:                                   If mdbeAdoptions.Errors(0).Number = daoERREUR_03200 Then mlngDernièreErreurDAO = daoENREGISTREMENT_CONNEXE_EXISTANT: Resume Next
                                                                If mdbeAdoptions.Errors(0).Number = daoERREUR_03022 Then mlngDernièreErreurDAO = daoDOUBLON: Resume Next
                                                                If mdbeAdoptions.Errors(0).Number = daoERREUR_03201 Then mlngDernièreErreurDAO = daoENREGISTREMENT_CONNEXE_MANQUANT: Resume Next
                                                                If mdbeAdoptions.Errors(0).Number = daoERREUR_03035 Or mdbeAdoptions.Errors(0).Number = daoERREUR_03260 Or mdbeAdoptions.Errors(0).Number = daoERREUR_03261 Then pbooVerrouillageActivé = True: Resume Next
                                                                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
        Case mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrMessages(0), pstrMessages(1))
        Case mesANNULATION_SUR_DOUBLON:                         pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, "La mise à jour", "", pstrChaîneSQL)
        Case mesANNULATION_SUR_CONNEXE_MANQUANT:                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, "La mise à jour", "", pstrNomTable, pstrChaîneSQL)
        Case mesANNULATION_SUR_CONNEXE_EXISTANT:                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, "mise à jour", pstrNomTable, pstrChaîneSQL)
        Case mesMESSAGE_AVEC_ICÔNE:                             pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrMessages(0), pstrMessages(1))
        Case Else:                                              pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub


'Supprimer un ou plusieurs enregistrements
'   Les valeurs prises par mlngDernièreErreurDAO sont:  (*) daoENREGISTREMENT_VERROUILLÉ:       Les données sont verrouillées par une transaction
'                                                           daoENREGISTREMENT_CONNEXE_EXISTANT: La suppression a été refusée du fait d'un enregistrement connexe existant
'                                                       (*) daoAUCUNE:                          L'(Les )enregistrement(s) a(ont) été supprimé(s)
'       (*) la valeur est retournée et peut donc être interceptée par la procédure appelante
'   plngNuméroRequête:  numéro de la requête
'   pbooGérerSablier:   booléen précisant s'il faut gérer le sablier
'   pstrMessage:        message à afficher en cas de verrouillage au format '<message si verrouillage[¤pictogramme]>'
'                       3 formats possibles:    "":                         aucun message n'est affiché
'                                               "<message>":                le message est affiché avec le pictogramme vbExclamation
'                                               "<message>¤<pictogramme>":  le message est affiché avec le pictogramme désigné
'   pvarArgument():     tableau d'arguments

Public Sub SQLSupprimer(plngNuméroRequête As DAOSuppressions, pbooGérerSablier As Boolean, ByVal pstrMessage As String, ParamArray pvarArgument() As Variant)
Const pkstrNomProcédure As String = "SQLSupprimer"
Dim pdatHeureRéférence As Date
Dim pstrChaîneSQL As String, pstrArguments() As String, pstrMessages() As String, pstrNomTable As String
Dim pintIndiceArgument As Integer, pintPositionApostrophe() As Integer, pintNuméroApostrophe() As Integer, pintIndiceTableau As Integer, pintPositionDépartRecherche As Integer
Dim pbooVerrouillageActivé As Boolean
On Error GoTo Err_SQLSupprimer

    If pbooGérerSablier Then Screen.MousePointer = vbHourglass
    
    If Not UBound(pvarArgument) = comINDÉFINI Then ReDim pstrArguments(LBound(pvarArgument) To UBound(pvarArgument))
    
'   Pour chaque argument passé en référence
    For pintIndiceArgument = LBound(pvarArgument) To UBound(pvarArgument)
        Select Case VarType(pvarArgument(pintIndiceArgument))
            Case vbBoolean
                pstrArguments(pintIndiceArgument) = "= " & IIf(pvarArgument(pintIndiceArgument), "TRUE", "FALSE")
            Case vbDate
                pstrArguments(pintIndiceArgument) = "= " & strSQLDateFormatée(pvarArgument(pintIndiceArgument))
            Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                pstrArguments(pintIndiceArgument) = "= " & strSQLNombreFormaté(pvarArgument(pintIndiceArgument))
            Case vbNull
                pstrArguments(pintIndiceArgument) = "IS NULL"
            Case vbString
                pstrArguments(pintIndiceArgument) = "= '" & strSQLChaîneFormatée(pvarArgument(pintIndiceArgument)) & "'"
        End Select
    Next pintIndiceArgument

    Select Case plngNuméroRequête

'       Suppression de l'historique d'une commission, suppression du recours administratif ou des recours contentieux lié(s) à une décision de commission
        Case daoDEL_HISTORIQUE_COMMISSION_OU_RECOURS
            pstrChaîneSQL = "DELETE " _
                          & "FROM " & pvarArgument(0) & " " _
                          & "WHERE (" & pvarArgument(0) & ".DEMA_NUM " & pstrArguments(1) & " AND " & pvarArgument(0) & ".u02_dtmHistoCommissionReunion " & pstrArguments(2) & ")"

'       Suppression d'un enfant/de ses origines souhaitées/de sa demande/de son accord en commission
        Case daoDEL_ORIGINES_OU_ENFANT_OU_DEMANDÉ_OU_ACCORDÉ
            pstrChaîneSQL = "DELETE " _
                          & "FROM " & pvarArgument(0) & " " _
                          & "WHERE (" & pvarArgument(0) & ".ENDE_NUM " & pstrArguments(1) & ")"

'       Suppression de l'historique d'un intervenant
        Case daoDEL_HISTORIQUE_INTERVENANT
            pstrChaîneSQL = "DELETE " _
                          & "FROM u01_HistoIntervenants " _
                          & "WHERE (u01_HistoIntervenants.DEMA_NUM " & pstrArguments(0) & " AND u01_HistoIntervenants.u01_booHistoIntervenantPsychologue " & pstrArguments(1) & " AND u01_HistoIntervenants.u01_dtmHistoIntervenantDemande " & pstrArguments(2) & " AND u01_HistoIntervenants.u01_intHistoIntervenantNumero " & pstrArguments(3) & ")"

'       frm_PlacementGestion: Suppression d'un placement/des enfants liés à un placement
        Case daoSUPPRESSION_001
            pstrChaîneSQL = "DELETE " _
                          & "FROM " & pvarArgument(0) & " " _
                          & "WHERE (" & pvarArgument(0) & ".PLAC_NUM " & pstrArguments(1) & ")"

    End Select

    pdatHeureRéférence = Now
    
    Do
    
'       Si les données ont été verrouillées par une transaction
        If pbooVerrouillageActivé Then
'           Avertir l'utilisateur du verrouillage des données au cas où le délai maximal avant l'affichage d'un message serait atteint
            GérerVerrouillage pbooVerrouillageActivé, pdatHeureRéférence
'           Annuler l'opération si l'utilisateur le demande
            If mlngDernièreErreurDAO = daoENREGISTREMENT_VERROUILLÉ Then
'               Si un message supplémentaire doit être affiché
                If Not pstrMessage = "" Then
                    pstrMessages() = Split(pstrMessage, gstrSÉPARATEUR_2)
'                   Si aucun pictogramme n'est défini
                    If UBound(pstrMessages) = 0 Then
                        ReDim Preserve pstrMessages(1)
                        pstrMessages(1) = vbExclamation
                    End If
                    Err.Raise mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION
                End If
                If pbooGérerSablier Then Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        
        mlngDernièreErreurDAO = daoAUCUNE
        
'       Exécuter la suppression
        mconnODBCAdoptions.Execute pstrChaîneSQL, dbExecDirect
        DoEvents
        
'   Boucler tant qu'une transaction verrouille les données
    Loop While pbooVerrouillageActivé
    
'   Si la suppression a été refusée du fait d'un enregistrement connexe existant
    If mlngDernièreErreurDAO = daoENREGISTREMENT_CONNEXE_EXISTANT Then
        ReDim pintPositionApostrophe(1 To ERREUR_3200_NOMBRE_APOSTROPHES) As Integer: ReDim pintNuméroApostrophe(0) As Integer: pintNuméroApostrophe(0) = CInt(ERREUR_3200_NUMÉRO_APOSTROPHE_AVANT_NOM_TABLE)
'       Extraire les positions des apostrophes dans le message
        For pintIndiceTableau = 1 To UBound(pintPositionApostrophe)
            If pintIndiceTableau = 1 Then pintPositionDépartRecherche = 1 Else pintPositionDépartRecherche = pintPositionApostrophe(pintIndiceTableau - 1) + 1
            pintPositionApostrophe(pintIndiceTableau) = InStr(pintPositionDépartRecherche, mdbeAdoptions.Errors(0).Description, "'")
        Next pintIndiceTableau
'       Extraire du message d'erreur DAO le nom de la table connexe
        pstrNomTable = Mid(mdbeAdoptions.Errors(0).Description, pintPositionApostrophe(pintNuméroApostrophe(0)) + 1, pintPositionApostrophe(pintNuméroApostrophe(0) + 1) - pintPositionApostrophe(pintNuméroApostrophe(0)) - 1)
        Err.Raise mesANNULATION_SUR_CONNEXE_EXISTANT
    End If
                        
    If pbooGérerSablier Then Screen.MousePointer = vbDefault
    
    Exit Sub

Err_SQLSupprimer:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                                     pintTraitement = mesINTERROMPRE
        Case daoERREUR_03146:                                   If mdbeAdoptions.Errors(0).Number = daoERREUR_03200 Then mlngDernièreErreurDAO = daoENREGISTREMENT_CONNEXE_EXISTANT: Resume Next
                                                                If mdbeAdoptions.Errors(0).Number = daoERREUR_03049 Or mdbeAdoptions.Errors(0).Number = daoERREUR_03260 Or mdbeAdoptions.Errors(0).Number = daoERREUR_03261 Then pbooVerrouillageActivé = True: Resume Next
                                                                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
        Case mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrMessages(0), pstrMessages(1))
        Case mesANNULATION_SUR_CONNEXE_EXISTANT:                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, "suppression", pstrNomTable, pstrChaîneSQL)
        Case Else:                                              pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub


'Retourner une date après l'avoir formatée en vue de son utilisation dans du code SQL
'   pdtmDate:   date dont on souhaite le formatage

Public Function strSQLDateFormatée(ByVal pdtmDate As Date) As String
Const pkstrNomProcédure As String = "strSQLDateFormatée"
On Error GoTo Err_strSQLDateFormatée

    strSQLDateFormatée = "#" & Format(pdtmDate, gstrFORMAT_DATE_AMÉRICAIN) & "#"

    Exit Function
    
Err_strSQLDateFormatée:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Function


'Retourner la chaîne de connexion ODBC

Public Property Get strConnexionODBC() As String
Const pkstrNomProcédure As String = "strConnexionODBC"
On Error GoTo Err_strConnexionODBC

    If Not mconnODBCAdoptions Is Nothing Then strConnexionODBC = mconnODBCAdoptions.Connect
    
    Exit Property
    
Err_strConnexionODBC:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Property


'Retourner le nom de la base avec son chemin

Public Property Get strNomBaseAvecChemin() As String
Const pkstrNomProcédure As String = "strNomBaseAvecChemin"
On Error GoTo Err_strNomBaseAvecChemin

    strNomBaseAvecChemin = mstrNomBaseAvecChemin
    
    Exit Property
    
Err_strNomBaseAvecChemin:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Property


'Retourner une chaîne après avoir doublé ses apostrophes en vue de son utilisation dans du code SQL
'   L'introduction dans du code SQL d'une chaîne contenant une apostrophe simple pose problème: chaque apostrophe doit donc être doublée
'   pstrChaîne: chaîne dont on souhaite le formatage

Public Function strSQLChaîneFormatée(ByVal pstrChaîne As String) As String
Const pkstrNomProcédure As String = "strSQLChaîneFormatée"
On Error GoTo Err_strSQLChaîneFormatée

    strSQLChaîneFormatée = Replace(pstrChaîne, "'", "''")
    
    Exit Function
    
Err_strSQLChaîneFormatée:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Function


'Retourner un nombre après avoir l'avoir formaté en vue de son utilisation dans du code SQL
'   pvarNombre: nombre dont on souhaite le formatage

Public Function strSQLNombreFormaté(pvarNombre As Variant) As String
Const pkstrNomProcédure As String = "strSQLNombreFormaté"
Dim pintPositionCaractère As Integer
Dim pstrNombre As String
Dim pstrChaine As String
On Error GoTo Err_strSQLNombreFormaté

    strSQLNombreFormaté = Replace(CStr(pvarNombre), ",", ".")
    
    Exit Function

Err_strSQLNombreFormaté:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Function


'Supprimer les classes liées aux données

Public Sub SupprimerClasses()
Const pkstrNomProcédure As String = "SupprimerClasses"
On Error GoTo Err_SupprimerClasses

    Set gclsProvenances = Nothing: Set gclsRecoursContentieux = Nothing

    Exit Sub

Err_SupprimerClasses:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Valider une opération DAO sur un jeu d'enregistrements
'   Les valeurs prises par mlngDernièreErreurDAO sont:  (*) daoENREGISTREMENT_VERROUILLÉ:       Les données sont verrouillées par une transaction
'                                                       (#) daoDOUBLON:                         L'opération a été refusée pour éviter la création d'un doublon
'                                                           daoENREGISTREMENT_CONNEXE_MANQUANT: L'opération a été refusée du fait d'un enregistrement connexe manquant
'                                                           daoENREGISTREMENT_CONNEXE_EXISTANT: L'opération a été refusée du fait d'un enregistrement connexe existant
'                                                       (*) daoAUCUNE:                          L'enregistrement a été supprimé (ou modifié)
'                                                       (*) dbRecordModified:                   Les données ont été modifiées entre-temps par un autre utilisateur
'                                                       (¤) dbRecordDeleted:                    L'enregistrement a été supprimé entre-temps par un autre utilisateur
'                                                       (*) prstCible.RecordStatus:             Conflit d'accès inattendu
'                                                       (*) daoERREUR_01809:                    La création a été refusée du fait de l'ouverture, par un autre utilisateur, de la base ou de l'objet en mode lecture seule
'       (*) la valeur est retournée et peut donc être interceptée par la procédure appelante
'       (#) la valeur n'est retournée que si un message est défini (message si doublon); si le message correspond à gstrDOUBLON_SANS_MESSAGE, la valeur est retourné sans aucun message
'       (¤) la valeur n'est pas retournée s'il s'agit d'une suppression
'   prstCible:                      jeu d'enregistrements sur lequel porte l'opération
'   pbooSupprimer:                  booléen précisant s'il s'agit de supprimer un enregistrement [facultatif]
'   pstrListeMessages:              liste de messages déterminant les messages supplémentaires à afficher en cas d'erreur au format '<message si verrouillage[¤pictogramme]>[|<message si suppression[¤pictogramme]>[|<message si modification et choix de ne pas écraser[¤pictogramme]>[|<message si doublon[¤pictogramme]>]]]' [facultatif]
'                                   pour chaque message, 3 formats possibles:   "":                         aucun message n'est affiché
'                                                                               "<message>":                le message est affiché avec le pictogramme vbExclamation
'                                                                               "<message>¤<pictogramme>":  le message est affiché avec le pictogramme désigné
'   pvarRéponseÀInviteÉcrasement:   constante précisant le choix d'écraser ou non les modifications éventuellement apportées entre-temps par un autre utilisateur, ou de laisser les données en l'état [facultatif: si elle est manquante, le choix d'écraser ou de ne pas écraser sera proposé à l'utilisateur]
'                                   valeurs possibles: vbYes, vbNo, vbCancel

Public Sub ValiderOpérationSurRecordset(prstCible As Recordset, Optional pbooSupprimer As Boolean = False, Optional ByVal pstrListeMessages As String = gstrMESSAGE_SUR_VERROUILLAGE_EN_VALIDATION & gstrSÉPARATEUR_1 & gstrMESSAGE_SUR_SUPPRESSION & gstrSÉPARATEUR_1 & gstrMESSAGE_SUR_MODIFICATION_ET_CHOIX_ANNULER, Optional pvarRéponseÀInviteÉcrasement As Variant)
Const pkstrNomProcédure As String = "ValiderOpérationSurRecordset"
Const pkintVerrouillage As Integer = 0, pkintSuppression As Integer = 1, pkintModification As Integer = 2, pkintDoublon As Integer = 3
Dim pfldÉlément As Field
Dim pvarValeurAutreUtilisateur As Variant, pvarColonnes() As Variant
Dim pdatHeureRéférence As Date
Dim pstrListesMessages() As String, pstrMessages() As String, pstrNomTable As String
Dim plngStatus As DAO.RecordStatusEnum, plngEditMode As DAO.EditModeEnum
Dim pintPositionApostrophe() As Integer, pintNuméroApostrophe() As Integer, pintIndiceTableau As Integer, pintPositionDépartRecherche As Integer, pintItération As Integer
Dim pbooVerrouillageActivé As Boolean
On Error GoTo Err_ValiderOpérationSurRecordset

    pdatHeureRéférence = Now
    
    pstrListesMessages() = Split(pstrListeMessages, gstrSÉPARATEUR_1)
    While Not UBound(pstrListesMessages) = pkintDoublon
        pstrListeMessages = pstrListeMessages & gstrSÉPARATEUR_1
        pstrListesMessages() = Split(pstrListeMessages, gstrSÉPARATEUR_1)
    Wend
    
    With prstCible
    
        plngStatus = .RecordStatus
        
#If Not TablesEnLocalOuAttachées Then
        plngEditMode = .EditMode

'       S'il ne s'agit pas d'une suppression
        If Not pbooSupprimer Then
            ReDim pvarColonnes(.Fields.Count - 1)
'           Sauvegarder les valeurs en cours de l'enregistrement
            For pintItération = 0 To UBound(pvarColonnes, 1)
                pvarColonnes(pintItération) = .Fields(pintItération).Value
            Next pintItération
        End If
#End If
        
        Do
        
'           Si les données ont été verrouillées par une transaction
            If pbooVerrouillageActivé Then

#If Not TablesEnLocalOuAttachées Then
'               S'il s'agit d'un ajout
                If plngEditMode = dbEditAdd Then
                    mlngDernièreErreurDAO = daoERREUR_01809
                    Err.Raise mesMESSAGE_023
                    Exit Sub
                End If
#End If

'               Avertir l'utilisateur du verrouillage au cas où le délai maximal avant l'affichage d'un message serait atteint
                GérerVerrouillage pbooVerrouillageActivé, pdatHeureRéférence
'               Annuler l'opération si l'utilisateur le demande
                If mlngDernièreErreurDAO = daoENREGISTREMENT_VERROUILLÉ Then
'                   Si un message supplémentaire doit être affiché
                    If Not pstrListesMessages(pkintVerrouillage) = "" Then
                        pstrMessages() = Split(pstrListesMessages(pkintVerrouillage), gstrSÉPARATEUR_2)
'                       Si aucun pictogramme n'est défini
                        If UBound(pstrMessages) = 0 Then
                            ReDim Preserve pstrMessages(1)
                            pstrMessages(1) = vbExclamation
                        End If
                        Err.Raise mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION
                    End If
                    Exit Sub
                End If
                
#If Not TablesEnLocalOuAttachées Then
'               Se repositionner sur l'enregistrement
                .Move 0
                
'               S'il ne s'agit pas d'une suppression
                If Not pbooSupprimer Then
                    .Edit
'                   Recharger l'enregistrement à partir des valeurs sauvegardées
                    For pintItération = 0 To UBound(pvarColonnes)
'                       Si la valeur de la colonne a été modifiée, la recharger
'                           Notes:
'                               Le fait de conditionner le rechargement évite de générer l'erreur daoERREUR_03200 lorsque l'on charge une clé (même si la valeur reste identique)
'                               La valeur Null se propage (ex: <1 = Null> donne Null, de même que <Null OR False>) mais <Null Or True> donne True
                        If (Not .Fields(pintItération).Value = pvarColonnes(pintItération)) Or (Not IsNull(.Fields(pintItération).Value) And IsNull(pvarColonnes(pintItération))) Or (IsNull(.Fields(pintItération).Value) And Not IsNull(pvarColonnes(pintItération))) Then .Fields(pintItération).Value = pvarColonnes(pintItération)
                    Next pintItération
                End If
#End If

            End If
            
            mlngDernièreErreurDAO = daoAUCUNE
    
'           Valider l'opération
#If TablesEnLocalOuAttachées Then
'               Notes:
'                   Si l'opération réussit, le statut de l'enregistrement (.RecordStatus):
'                       - pour un ajout:        passe de dbRecordNew (2) à dbRecordUnmodified (0)
'                       - pour une mise à jour: passe de dbRecordModified (1) à dbRecordUnmodified (0)
'                       - pour une suppression: reste à dbRecordDeleted (3)
'                   Si l'opération échoue, le statut de l'enregistrement (.RecordStatus) n'est pas modifié mais:
'                       - le nombre d'enregistrements ayant généré une collision (.BatchCollisionCount) passe à 1, sauf si l'opération a été refusée pour éviter la création d'un doublon
            If pbooSupprimer Then .Delete
            .Update dbUpdateCurrentRecord
#Else
'               Notes:
'                   Si l'opération réussit:
'                       - le statut de l'enregistrement (.RecordStatus):
'                           - pour un ajout:        passe de dbRecordNew (2) à dbRecordUnmodified (0)
'                           - pour une mise à jour: reste à dbRecordUnmodified (0)
'                           - pour une suppression: passe dbRecordUnmodified (0) à de dbRecordDBDeleted (4)
'                       - l'état de modification (.EditMode):
'                           - pour un ajout:        passe de dbEditAdd (2) à dbEditNone (0)
'                           - pour une mise à jour: passe de dbEditInProgress (1) à dbEditNone (0)
'                           - pour une suppression: reste à dbEditNone (0)
            If pbooSupprimer Then .Delete Else .Update
#End If
            DoEvents
            
'       Boucler tant qu'une transaction verrouille les données
        Loop While pbooVerrouillageActivé
        
'       Si l'opération a été refusée pour éviter la création d'un doublon
        If mlngDernièreErreurDAO = daoDOUBLON Then
        
'           Si l'erreur n'est pas gérée
            If pstrListesMessages(pkintDoublon) = "" Then
                Err.Raise mesANNULATION_SUR_DOUBLON
'           Sinon
            Else
'               Si l'erreur est gérée après l'affichage d'un message
                If Not pstrListesMessages(pkintDoublon) = gstrDOUBLON_SANS_MESSAGE Then
                    pstrMessages() = Split(pstrListesMessages(pkintDoublon), gstrSÉPARATEUR_2)
'                   Si aucun pictogramme n'est défini
                    If UBound(pstrMessages) = 0 Then
                        ReDim Preserve pstrMessages(1)
                        pstrMessages(1) = vbExclamation
                    End If
                    Err.Raise mesMESSAGE_AVEC_ICÔNE
                End If
            End If
                
'       Sinon, si l'opération a été refusée du fait d'un enregistrement connexe manquant
        ElseIf mlngDernièreErreurDAO = daoENREGISTREMENT_CONNEXE_MANQUANT Then

            ReDim pintPositionApostrophe(1 To ERREUR_3201_NOMBRE_APOSTROPHES) As Integer: ReDim pintNuméroApostrophe(0) As Integer: pintNuméroApostrophe(0) = CInt(ERREUR_3201_NUMÉRO_APOSTROPHE_AVANT_NOM_TABLE)
'           Extraire les positions des apostrophes dans le message
            For pintIndiceTableau = 1 To UBound(pintPositionApostrophe)
                If pintIndiceTableau = 1 Then pintPositionDépartRecherche = 1 Else pintPositionDépartRecherche = pintPositionApostrophe(pintIndiceTableau - 1) + 1
                pintPositionApostrophe(pintIndiceTableau) = InStr(pintPositionDépartRecherche, mdbeAdoptions.Errors(0).Description, "'")
            Next pintIndiceTableau
'           Extraire du message d'erreur DAO le nom de la table connexe
            pstrNomTable = Mid(mdbeAdoptions.Errors(0).Description, pintPositionApostrophe(pintNuméroApostrophe(0)) + 1, pintPositionApostrophe(pintNuméroApostrophe(0) + 1) - pintPositionApostrophe(pintNuméroApostrophe(0)) - 1)
            Err.Raise mesANNULATION_SUR_CONNEXE_MANQUANT
                    
'       Sinon, si l'opération a été refusée du fait d'un enregistrement connexe existant
        ElseIf mlngDernièreErreurDAO = daoENREGISTREMENT_CONNEXE_EXISTANT Then

            ReDim pintPositionApostrophe(1 To ERREUR_3200_NOMBRE_APOSTROPHES) As Integer: ReDim pintNuméroApostrophe(0) As Integer: pintNuméroApostrophe(0) = CInt(ERREUR_3200_NUMÉRO_APOSTROPHE_AVANT_NOM_TABLE)
'           Extraire les positions des apostrophes dans le message
            For pintIndiceTableau = 1 To UBound(pintPositionApostrophe)
                If pintIndiceTableau = 1 Then pintPositionDépartRecherche = 1 Else pintPositionDépartRecherche = pintPositionApostrophe(pintIndiceTableau - 1) + 1
                pintPositionApostrophe(pintIndiceTableau) = InStr(pintPositionDépartRecherche, mdbeAdoptions.Errors(0).Description, "'")
            Next pintIndiceTableau
'           Extraire du message d'erreur DAO le nom de la table connexe
            pstrNomTable = Mid(mdbeAdoptions.Errors(0).Description, pintPositionApostrophe(pintNuméroApostrophe(0)) + 1, pintPositionApostrophe(pintNuméroApostrophe(0) + 1) - pintPositionApostrophe(pintNuméroApostrophe(0)) - 1)
            Err.Raise mesANNULATION_SUR_CONNEXE_EXISTANT
                
#If TablesEnLocalOuAttachées Then
'       Sinon, si l'opération a été refusée pour avoir entraîné une collision
        ElseIf .BatchCollisionCount = 1 Then
    
'           S'il s'agit de supprimer un enregistrement déjà supprimé, sortir
            If plngStatus = dbRecordDeleted And prstCible.RecordStatus = dbRecordDeleted Then Exit Sub
    
'           Lire la valeur du premier champ telle qu'elle est stockée dans la base
'               Note: Force le positionnement du statut de l'enregistrement et l'alimentation la collection Errors si celui-ci a été supprimé par un utilisateur
            pvarValeurAutreUtilisateur = .Fields(0).VisibleValue

            mlngDernièreErreurDAO = .RecordStatus

'           Si l'enregistrement a été supprimé par un utilisateur
            If .RecordStatus = dbRecordDeleted Then
'               Si un message supplémentaire doit être affiché
                If Not pstrListesMessages(pkintSuppression) = "" Then
                    pstrMessages() = Split(pstrListesMessages(pkintSuppression), gstrSÉPARATEUR_2)
'                   Si aucun pictogramme n'est défini
                    If UBound(pstrMessages) = 0 Then
                        ReDim Preserve pstrMessages(1)
                        pstrMessages(1) = vbExclamation
                    End If
                    Err.Raise mesMESSAGE_AVEC_ICÔNE
                End If
                Exit Sub
            End If

'           Si le choix d'écraser ou non les modifications apportées entre-temps par un autre utilisateur n'est pas défini, le proposer à l'utilisateur
            If IsMissing(pvarRéponseÀInviteÉcrasement) Then pvarRéponseÀInviteÉcrasement = modMessages.lngRéponseUtilisateur(mesQUESTION_002)

            Select Case pvarRéponseÀInviteÉcrasement
'               Si le choix est d'écraser les modifications
                Case vbYes
                    mlngDernièreErreurDAO = daoAUCUNE
'               Si le choix est de ne pas écraser les modifications
                Case vbNo
'                   Préparer l'actualisation de l'enregistrement
                    .Edit
'                   Recharger l'enregistrement à partir des valeurs modifiées par l'autre utilisateur
                    For Each pfldÉlément In .Fields
'                       Si la valeur de la colonne a été modifiée par l'autre utilisateur, la recharger
'                           Notes:
'                               Le fait de conditionner le rechargement évite de générer l'erreur daoERREUR_03200 lorsque l'on charge une clé (même si la valeur reste identique)
'                               La valeur Null se propage (ex: <1 = Null> donne Null, de même que <Null OR False>) mais <Null Or True> donne True
                        With pfldÉlément
                            If (Not .Value = .VisibleValue) Or (Not IsNull(.Value) And IsNull(.VisibleValue)) Or (IsNull(.Value) And Not IsNull(.VisibleValue)) Then .Value = .VisibleValue
                        End With
                    Next pfldÉlément
            End Select

'           Forcer la mise à jour de l'enregistrement
'               Note: Si le choix est d'écraser les modifications, les données sont mises à jour à partir des données saisies par l'utilisateur; sinon, elles le sont à partir des données modifiées par l'autre utilisateur
            .Update dbUpdateCurrentRecord, True
    
'           Si le choix est de ne pas écraser les modifications
            If pvarRéponseÀInviteÉcrasement = vbNo Then
'               Si un message supplémentaire doit être affiché
                If Not pstrListesMessages(pkintModification) = "" Then
                    pstrMessages() = Split(pstrListesMessages(pkintModification), gstrSÉPARATEUR_2)
'                   Si aucun pictogramme n'est défini
                    If UBound(pstrMessages) = 0 Then
                        ReDim Preserve pstrMessages(1)
                        pstrMessages(1) = vbExclamation
                    End If
                    Err.Raise mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION
                End If
            End If
#Else
'       Si l'enregistrement a été supprimé par un utilisateur alors qu'il ne s'agit pas de supprimer un enregistrement déjà supprimé
        ElseIf mlngDernièreErreurDAO = dbRecordDeleted Then
    
'           Si un message supplémentaire doit être affiché
            If Not pstrListesMessages(pkintSuppression) = "" Then
                pstrMessages() = Split(pstrListesMessages(pkintSuppression), gstrSÉPARATEUR_2)
'               Si aucun pictogramme n'est défini
                If UBound(pstrMessages) = 0 Then
                    ReDim Preserve pstrMessages(1)
                    pstrMessages(1) = vbExclamation
                End If
                Err.Raise mesMESSAGE_AVEC_ICÔNE
            End If
#End If
    
        End If
        
    End With
    
    Exit Sub

Err_ValiderOpérationSurRecordset:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:                                     pintTraitement = mesINTERROMPRE
        Case daoERREUR_03146:                                   If mdbeAdoptions.Errors(0).Number = daoERREUR_03022 Then mlngDernièreErreurDAO = daoDOUBLON: Resume Next
                                                                If mdbeAdoptions.Errors(0).Number = daoERREUR_03200 Then mlngDernièreErreurDAO = daoENREGISTREMENT_CONNEXE_EXISTANT: Resume Next
                                                                If mdbeAdoptions.Errors(0).Number = daoERREUR_03201 Then mlngDernièreErreurDAO = daoENREGISTREMENT_CONNEXE_MANQUANT: Resume Next
                                                                If mdbeAdoptions.Errors(0).Number = daoERREUR_01304 Or mdbeAdoptions.Errors(0).Number = daoERREUR_03035 Or mdbeAdoptions.Errors(0).Number = daoERREUR_03049 Or mdbeAdoptions.Errors(0).Number = daoERREUR_03260 Or mdbeAdoptions.Errors(0).Number = daoERREUR_03261 Or (plngStatus = dbRecordDeleted And prstCible.RecordStatus = dbRecordDeleted) Then pbooVerrouillageActivé = True: Resume Next
#If TablesEnLocalOuAttachées Then
                                                                If prstCible.RecordStatus = dbRecordDeleted Then Resume Next
#Else
                                                                If mdbeAdoptions.Errors(0).Number = daoERREUR_01809 Then pbooVerrouillageActivé = True: Resume Next
                                                                If mdbeAdoptions.Errors(0).Number = daoERREUR_01017 Then Resume Next
                                                                If mdbeAdoptions.Errors(0).Number = daoERREUR_01611 And prstCible.RecordStatus = dbRecordDBDeleted Then mlngDernièreErreurDAO = dbRecordDeleted: Resume Next
#End If
                                                                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
        Case mesANNULATION_SUR_VERROUILLAGE_OU_MODIFICATION:    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrMessages(0), pstrMessages(1))
        Case mesMESSAGE_AVEC_ICÔNE:                             pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, pstrMessages(0), pstrMessages(1))
        Case mesANNULATION_SUR_DOUBLON:                         pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, "La mise à jour", "e", prstCible.Name)
        Case mesANNULATION_SUR_CONNEXE_MANQUANT:                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, "La mise à jour", "e", pstrNomTable, prstCible.Name)
#If TablesEnLocalOuAttachées Then
        Case mesANNULATION_SUR_CONNEXE_EXISTANT:                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, IIf(prstCible.RecordStatus = dbRecordModified, "mise à jour", "suppression"), pstrNomTable, prstCible.Name)
#Else
        Case mesANNULATION_SUR_CONNEXE_EXISTANT:                pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing, IIf(prstCible.EditMode = dbEditInProgress, "mise à jour", "suppression"), pstrNomTable, prstCible.Name)
        Case mesMESSAGE_023:                                    pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
#End If
        Case Else:                                              pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Sub


'Retourner la valeur maximale prise par la colonne d'une table
'   Les valeurs prises par mlngDernièreErreurDAO sont:  (*) daoENREGISTREMENT_VERROUILLÉ:   Les données sont verrouillées par une transaction
'                                                       (*) daoAUCUNE:                      La lecture a retourné un jeu d'enregistrements
'       (*) la valeur est retournée et peut donc être interceptée par la procédure appelante
'   pstrTable:                          table à laquelle la colonne appartient
'   pstrColonne:                        colonne dont on souhaite connaître la valeur maximale
'   pstrMessageSiVerrouillage:          message à afficher si les données sont verrouillées par une transaction
'   pvarValeurSiAucunEnregistrement:    valeur à retourner si la table ne contient aucun enregistrement

Public Function varValeurMaximale(pstrTable As String, pstrColonne As String, pstrMessageSiVerrouillage As String, Optional pvarValeurSiAucunEnregistrement As Variant = 0) As Variant
Const pkstrNomProcédure As String = "varValeurMaximale"
Dim prstMaximum As Recordset
On Error GoTo Err_varValeurMaximale

'   Lire la valeur maximale
    Set prstMaximum = modDAO.rstSQLExtraction(daoMAXIMUM, False, pstrMessageSiVerrouillage, pstrTable, pstrColonne)
    If mlngDernièreErreurDAO = daoAUCUNE Then varValeurMaximale = IIf(IsNull(prstMaximum.Fields(0)), pvarValeurSiAucunEnregistrement, prstMaximum.Fields(0))

    Exit Function
    
Err_varValeurMaximale:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End

End Function

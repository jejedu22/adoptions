Attribute VB_Name = "modCommun"
Option Explicit

Private Const mstrNOM_MODULE As String = "modCommun"

Public Enum CommunErreurs       'Erreurs Visual Basic
    comERREUR_00005 = 5         'Argument ou appel de procédure incorrect
    comERREUR_00006 = 6         'Dépassement de capacité
    comERREUR_00013 = 13        'Type incompatible
    comERREUR_00053 = 53        'Fichier introuvable
    comERREUR_00068 = 68        'Périphérique non disponible
    comERREUR_00070 = 70        'Permission refusée
    comERREUR_00071 = 71        'Disque non prêt
    comERREUR_00076 = 76        'Chemin d'accès introuvable
    comERREUR_00094 = 94        'Utilisation incorrecte de Null
    comERREUR_00438 = 438       'Cet objet ne gère pas cette propriété ou méthode
    comERREUR_00457 = 457       'Cette clé est déjà associée à un élément de cette collection
    comERREUR_03265 = 3265      'Élément non trouvé dans cette collection
End Enum
Public Enum CommunTraitements   'Types de traitement
    comCRÉATION = 1
    comCONSULTATION
    comMODIFICATION
    comSUPPRESSION
End Enum
Public Enum CommunDivers        'Constantes diverses
    comINDÉFINI = -1
    comLARGEUR_ASCENCEUR = 275
End Enum

Public Const gstrSÉPARATEUR_1 As String = "|", gstrSÉPARATEUR_2 As String = "¤", gstrSÉPARATEUR_3 As String = "•"
Public Const gstrFORMAT_DATE_AMÉRICAIN As String = "mm/dd/yyyy", gstrFORMAT_DATE_FRANÇAIS As String = "dd/mm/yyyy"

Private Const mstrREGISTRE_APPLICATION As String = "Software\VB and VBA Program Settings\Adoptions"


'Activer ou désactiver un contrôle CheckBox, ComboBox, CommandButton, DTPicker, FileListBox, Frame, Label, ListBox, ListView, MaskEdBox, OptionButton, TextBox ou Toolbar
'   Control:        contrôle dont on souhaite implémenter l'état
'   pbooVerrouillé: booléen précisant l'état souhaité du contrôle
'   pvarArgument(): tableau de type ParamArray constitué d'arguments (un argument n, facultatif, devient obligatoire si l'argument n+1 doit être précisé)
'       ComboBox
'           pvarArgument(0):    booléen précisant si la police doit être affichée en gras [facultatif]
'       DTPicker
'           pvarArgument(0):    booléen précisant si la police doit être affichée en gras [facultatif]
'       TextBox
'           pvarArgument(0):    booléen précisant si la police doit être affichée en gras [facultatif]
'       Toolbar
'           pvarArgument(0):    index du bouton dont on souhaite implémenter l'état [facultatif: s'il est manquant, tous les boutons du contrôle Toolbar seront implémentés]
'           pvarArgument(1):    valeur implémentant l'info-bulle au cas où le bouton est inactivé [facultatif]
'           pvarArgument(2):    libellé de l'info-bulle au cas où le bouton est activé [facultatif]

Public Sub DéfinirVerrouillageContrôle(Control As Object, pbooVerrouillé As Boolean, ParamArray pvarArgument() As Variant)
Const pkstrNomProcédure As String = "DéfinirVerrouillageContrôle"
Dim pctlCible As Control
Dim pvarÉlément As Variant
On Error GoTo Err_DéfinirVerrouillageContrôle

    Set pctlCible = Control
    
    With pctlCible
'       S'il s'agit d'un contrôle CheckBox
        If TypeOf pctlCible Is CheckBox Then
            .Enabled = Not pbooVerrouillé
            .TabStop = Not pbooVerrouillé
'       S'il s'agit d'un contrôle ComboBox
        ElseIf TypeOf pctlCible Is ComboBox Then
            .BackColor = IIf(pbooVerrouillé, vbButtonFace, vbWindowBackground)
            .Enabled = Not pbooVerrouillé
            .ForeColor = IIf(pbooVerrouillé, vbGrayText, vbButtonText)
            .Locked = pbooVerrouillé
            .TabStop = Not pbooVerrouillé
            If Not UBound(pvarArgument) = comINDÉFINI Then .FontBold = pvarArgument(0)
'       S'il s'agit d'un contrôle CommandButton
        ElseIf TypeOf pctlCible Is CommandButton Then
            .Enabled = Not pbooVerrouillé
            .TabStop = Not pbooVerrouillé
'       S'il s'agit d'un contrôle DTPicker
        ElseIf TypeOf pctlCible Is DTPicker Then
            .Enabled = Not pbooVerrouillé
            .TabStop = Not pbooVerrouillé
            If Not UBound(pvarArgument) = comINDÉFINI Then .Font.Bold = pvarArgument(0)
'       S'il s'agit d'un contrôle FileListBox
        ElseIf TypeOf pctlCible Is FileListBox Then
            .BackColor = IIf(pbooVerrouillé, vbButtonFace, vbWindowBackground)
            .Enabled = Not pbooVerrouillé
'       S'il s'agit d'un contrôle Frame
        ElseIf TypeOf pctlCible Is Frame Then
            .Enabled = Not pbooVerrouillé
'       S'il s'agit d'un contrôle Label
        ElseIf TypeOf pctlCible Is Label Then
            .Enabled = Not pbooVerrouillé
'       S'il s'agit d'un contrôle ListBox
        ElseIf TypeOf pctlCible Is ListBox Then
            .BackColor = IIf(pbooVerrouillé, vbButtonFace, vbWindowBackground)
            .ForeColor = IIf(pbooVerrouillé, vbGrayText, vbButtonText)
            .Locked = pbooVerrouillé
            .TabStop = Not pbooVerrouillé
'       S'il s'agit d'un contrôle ListView
        ElseIf TypeOf pctlCible Is ListView Then
            .BackColor = IIf(pbooVerrouillé, vbButtonFace, vbWindowBackground)
            .Enabled = Not pbooVerrouillé
            .TabStop = Not pbooVerrouillé
''       S'il s'agit d'un contrôle MaskEdBox
'        ElseIf TypeOf pctlCible Is MaskEdBox Then
'            .BackColor = IIf(pbooVerrouillé, vbButtonFace, vbWindowBackground)
'            .Enabled = Not pbooVerrouillé
'            .TabStop = Not pbooVerrouillé
'       S'il s'agit d'un contrôle OptionButton
        ElseIf TypeOf pctlCible Is OptionButton Then
            .Enabled = Not pbooVerrouillé
            .TabStop = Not pbooVerrouillé
'       S'il s'agit d'un contrôle TextBox
        ElseIf TypeOf pctlCible Is textbox Then
            .BackColor = IIf(pbooVerrouillé, vbButtonFace, vbWindowBackground)
            .Enabled = Not pbooVerrouillé
            .ForeColor = IIf(pbooVerrouillé, vbGrayText, vbButtonText)
            .Locked = pbooVerrouillé
            .TabStop = Not pbooVerrouillé
            If Not UBound(pvarArgument) = comINDÉFINI Then .FontBold = pvarArgument(0)
'       S'il s'agit d'un contrôle Toolbar
        ElseIf TypeOf pctlCible Is Toolbar Then
            If UBound(pvarArgument) = comINDÉFINI Then
                For Each pvarÉlément In pctlCible.Buttons
                    If Not (pvarÉlément.Enabled = Not pbooVerrouillé) Then pvarÉlément.Enabled = Not pbooVerrouillé
                Next pvarÉlément
            Else
                With pctlCible.Buttons(pvarArgument(0))
                    If Not (.Enabled = Not pbooVerrouillé) Then .Enabled = Not pbooVerrouillé
                    If UBound(pvarArgument) = 2 Then .ToolTipText = IIf(pbooVerrouillé, pvarArgument(1) & " impossible", pvarArgument(2))
                End With
            End If
        End If
    End With
    
    Exit Sub
    
Err_DéfinirVerrouillageContrôle:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Forcer si nécessaire l'actualisation de la valeur d'un contrôle DTPicker
'   Notes:  L'actualisation de la propriété Value d'un DTPicker semble se faire avec un léger retard dans le cas où le focus reste sur le contrôle.
'           Un problème se pose quand le traitement interroge immédiatement la valeur sans que le focus n'ait été déplacé.
'           La solution consiste à forcer le déplacement du focus, puis à le replacer.
'   frmCible:   feuille faisant l'objet de l'actualisation
'   pctlFocus:  contrôle à qui donner le focus momentanément

Public Sub GérerActualisationDTPicker(frmCible As Form, pctlFocus As Control)
Const pkstrNomProcédure As String = "GérerActualisationDTPicker"
Dim pctlActif As Control
On Error GoTo Err_GérerActualisationDTPicker

'   Si le focus appartient à un contrôle DTPicker
    If TypeOf frmCible.ActiveControl Is DTPicker Then
        Set pctlActif = frmCible.ActiveControl
'       Le déplacer
        pctlFocus.SetFocus
'       Le replacer
        pctlActif.SetFocus
    End If
    
    Exit Sub

Err_GérerActualisationDTPicker:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Implémenter un contrôle ComboBox ou ListBox
'   pobjCible:  contrôle dont on souhaite l'implémentation
'   pvarIndex:  entier précisant la position du nouvel élément ou de la nouvelle ligne [facultatif]
'   pvarNuméro: numéro associé à l'élément [facultatif]

Public Sub ImplémenterListe(pobjCible As Object, pstrValeur As String, Optional pvarIndex As Variant, Optional pvarNuméro As Variant)
Const pkstrNomProcédure As String = "ImplémenterListe"
On Error GoTo Err_ImplémenterListe

    If Not IsMissing(pvarIndex) Then pobjCible.AddItem pstrValeur, CInt(pvarIndex) Else pobjCible.AddItem pstrValeur
    If Not IsMissing(pvarNuméro) Then pobjCible.ItemData(pobjCible.NewIndex) = CLng(pvarNuméro)
    
    Exit Sub

Err_ImplémenterListe:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Implémenter un contrôle ComboBox ou ListBox dont les éléments sont associés aux valeurs True/False
'   pobjCible:  contrôle dont on souhaite l'implémentation
'   pvarOuiNon: chaîne au format '<valeur Oui>,<valeur Non>' précisant les valeurs affichées [facultatif: si elle est manquante, la chaîne sera 'Oui,Non']

Public Sub ImplémenterListeOuiNon(pobjCible As Object, Optional pvarOuiNon As Variant)
Const pkstrNomProcédure As String = "ImplémenterListeOuiNon"
Dim pstrValeurs() As String
On Error GoTo Err_ImplémenterListeOuiNon

    If IsMissing(pvarOuiNon) Then
        ReDim pstrValeurs(1)
        pstrValeurs(0) = "Oui": pstrValeurs(1) = "Non"
    Else
        pstrValeurs = Split(pvarOuiNon, ",")
    End If
    
    ImplémenterListe pobjCible, pstrValeurs(0), pvarNuméro:=CLng(True)
    ImplémenterListe pobjCible, pstrValeurs(1), pvarNuméro:=CLng(False)
    
    Exit Sub

Err_ImplémenterListeOuiNon:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Implémenter les sous-éléments d'un objet ListItem
'   pitmCible:      objet ListItem dont on souhaite implémenter les sous-éléments
'   pvarArgument(): tableau de type ParamArray constitué de chaque valeur implémentant les sous-éléments de l'objet ListItem

Public Sub ImplémenterSubItems(pitmCible As ListItem, ParamArray pvarArgument() As Variant)
Const pkstrNomProcédure As String = "ImplémenterSubItems"
Dim pintIndiceTableau As Integer
On Error GoTo Err_ImplémenterSubItems

    For pintIndiceTableau = LBound(pvarArgument) To UBound(pvarArgument)
        pitmCible.SubItems(pintIndiceTableau + 1) = IIf(IsNull(pvarArgument(pintIndiceTableau)), "", pvarArgument(pintIndiceTableau))
    Next pintIndiceTableau

    Exit Sub
    
Err_ImplémenterSubItems:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Retourner un entier long précisant l'état d'une écriture sur le registre
'
'   plngCléStandard:        index d'un nom de clé standard
'   pstrCléCible:           chaîne de la clé sur laquelle porte l'écriture
'   pstrSousCléCible:       chaîne de la sous-clé sur laquelle porte l'écriture
'   plngTypeDonnées:        type de données du registre
'   pvarValeur:             valeur de la sous-clé dont on souhaite l'écriture
'   pbooCréerSiInexistante: booléen précisant si la création de la clé est exigée si elle n'existe pas dans le registre [facultatif]
'
'   Valeurs retournées: apiERROR_SUCCESS:           l'écriture s'est effectuée correctement
'                       apiERROR_FILE_NOT_FOUND:    la clé n'existe pas dans le registre et sa création n'est pas exigée
'                       toute autre valeur:         l'écriture a généré une erreur inattendue

Public Function lngÉcritureValeurRegistre(plngCléStandard As APIClésRacines, pstrCléCible As String, pstrSousCléCible As String, plngTypeDonnées As APITypeDonnées, pvarValeur As Variant, Optional pbooCréerSiInexistante As Boolean = False) As Long
Const pkstrNomProcédure As String = "lngÉcritureValeurRegistre"
Dim pstrValeur As String
Dim plngRetour As Long, plngHandle As Long, plngÉtat As Long
On Error GoTo Err_lngÉcritureValeurRegistre

'   Si la création de la clé est exigée
    If pbooCréerSiInexistante Then
    
'       Créer la clé; si elle existe déjà, l'ouvrir
        plngRetour = modAPI.RegCreateKeyEx(plngCléStandard, pstrCléCible, 0, plngTypeDonnées, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, plngHandle, plngÉtat)
        If Not plngRetour = apiERROR_SUCCESS Then lngÉcritureValeurRegistre = plngRetour: plngRetour = modAPI.RegCloseKey(plngHandle): Exit Function
        
    Else
    
'       Ouvrir la clé
        plngRetour = modAPI.RegOpenKeyEx(plngCléStandard, pstrCléCible, 0, KEY_ALL_ACCESS, plngHandle)
        If Not plngRetour = apiERROR_SUCCESS Then lngÉcritureValeurRegistre = plngRetour: plngRetour = modAPI.RegCloseKey(plngHandle): Exit Function
    
    End If
            
    Select Case plngTypeDonnées
    
'       S'il s'agit d'une chaîne avec un caractère NULL de terminaison
        Case apiREG_SZ
        
            pstrValeur = pvarValeur
'           Écrire la sous-clé et sa valeur
            plngRetour = modAPI.RegSetValueEx(plngHandle, pstrSousCléCible, 0, plngTypeDonnées, ByVal pstrValeur, CLng(Len(pstrValeur) + 1))
            If Not plngRetour = apiERROR_SUCCESS Then lngÉcritureValeurRegistre = plngRetour: plngRetour = modAPI.RegCloseKey(plngHandle): Exit Function
            
'       S'il s'agit d'un mot long de 32 bits
        Case apiREG_DWORD
        
'           Écrire la sous-clé et sa valeur
            plngRetour = modAPI.RegSetValueEx(plngHandle, pstrSousCléCible, 0, plngTypeDonnées, CLng(pvarValeur), 4)
            If Not plngRetour = apiERROR_SUCCESS Then lngÉcritureValeurRegistre = plngRetour: plngRetour = modAPI.RegCloseKey(plngHandle): Exit Function
            
    End Select
    
'   Refermer la clé
    lngÉcritureValeurRegistre = modAPI.RegCloseKey(plngHandle)
    
    Exit Function
    
Err_lngÉcritureValeurRegistre:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
            
End Function


'Retourner un entier long précisant l'état d'une lecture sur le registre
'
'   plngCléRacine:          index d'un nom de clé standard
'   plngCléSource:          chaîne de la clé sur laquelle porte la lecture
'   pstrSousCléSource:      chaîne de la sous-clé sur laquelle porte la lecture
'   plngTypeDonnées:        type de données du registre
'   pvarValeur:             variable recevant la valeur de la sous-clé après sa lecture
'   pbooCréerSiInexistante: booléen précisant si la création de la clé ou de la sous-clé est exigée si l'une ou l'autre n'existe pas dans le registre [facultatif]
'   pvarDéfaut:             valeur par défaut de la sous-clé dans le cas de sa création [obligatoire si une création est exigée]
'
'   Valeurs retournées: apiERROR_SUCCESS:           la lecture s'est effectuée correctement
'                       apiERROR_FILE_NOT_FOUND:    la clé ou la sous-clé n'existent pas dans le registre et leur création n'est pas exigée
'                       toute autre valeur:         la lecture a généré une erreur inattendue

Public Function lngLectureValeurRegistre(plngCléRacine As APIClésRacines, plngCléSource As String, pstrSousCléSource As String, plngTypeDonnées As APITypeDonnées, pvarValeur As Variant, Optional pbooCréerSiInexistante As Boolean = False, Optional pvarDéfaut As Variant) As Long
Const pkstrNomProcédure As String = "lngLectureValeurRegistre"
Dim pvarTampon As Variant
Dim pstrValeur As String, pstrValeurHexadécimale As String
Dim plngRetour As Long, plngHandle As Long, plngNombreOctets As Long, plngIndiceOctet As Long, plngValeur As Long, plngÉtat As Long
Dim pbooCléSousCléÀCréer As Boolean
On Error GoTo Err_lngLectureValeurRegistre

'   Si la création de la clé est exigée
    If pbooCréerSiInexistante Then
    
'       Créer la clé; si elle existe déjà, l'ouvrir
        plngRetour = modAPI.RegCreateKeyEx(plngCléRacine, plngCléSource, 0, plngTypeDonnées, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, plngHandle, plngÉtat)
        If Not plngRetour = apiERROR_SUCCESS Then lngLectureValeurRegistre = plngRetour: plngRetour = modAPI.RegCloseKey(plngHandle): Exit Function
        
    Else
    
'       Ouvrir la clé
        plngRetour = modAPI.RegOpenKeyEx(plngCléRacine, plngCléSource, 0, KEY_ALL_ACCESS, plngHandle)
        If Not plngRetour = apiERROR_SUCCESS Then lngLectureValeurRegistre = plngRetour: plngRetour = modAPI.RegCloseKey(plngHandle): Exit Function
    
    End If
            
'   Si la clé a dû être créée
    If plngÉtat = REG_CREATED_NEW_KEY Then
        
        pbooCléSousCléÀCréer = True
    
    Else
    
        Select Case plngTypeDonnées
        
'           S'il s'agit d'une chaîne avec un caractère NULL de terminaison
            Case apiREG_SZ
            
'               Lire le nombre d'octets occupés par la valeur de la sous-clé
                plngRetour = modAPI.RegQueryValueEx(plngHandle, pstrSousCléSource, 0, (plngTypeDonnées), pvarTampon, plngNombreOctets)
                Select Case plngRetour
'                   Si la sous-clé n'existe pas
                    Case apiERROR_FILE_NOT_FOUND
'                       Si la création de la sous-clé est exigée
                        If pbooCréerSiInexistante Then
                            pbooCléSousCléÀCréer = True
                        Else
'                           S'il existe une valeur par défaut pour la sous-clé
                            If Not IsMissing(pvarDéfaut) Then
                                pvarValeur = pvarDéfaut
                            Else
                                lngLectureValeurRegistre = plngRetour
                                plngRetour = modAPI.RegCloseKey(plngHandle)
                                Exit Function
                            End If
                        End If
'                   Si l'opération nécessite une deuxième lecture
                    Case apiERROR_MORE_DATA
                        pstrValeur = String(plngNombreOctets + 1, 0)
'                       Lire la valeur de la sous-clé
                        plngRetour = modAPI.RegQueryValueEx(plngHandle, pstrSousCléSource, 0, plngTypeDonnées, ByVal pstrValeur, plngNombreOctets)
                        If Not plngRetour = apiERROR_SUCCESS Then lngLectureValeurRegistre = plngRetour: plngRetour = modAPI.RegCloseKey(plngHandle): Exit Function
                        pstrValeur = Left(pstrValeur, plngNombreOctets - 1)
                        pvarValeur = pstrValeur
                    Case Else
                        lngLectureValeurRegistre = plngRetour
                        plngRetour = modAPI.RegCloseKey(plngHandle)
                        Exit Function
                End Select
                
'           S'il s'agit d'un mot long de 32 bits
            Case apiREG_DWORD
            
'               Lire le nombre d'octets occupés par la valeur de la sous-clé
                plngRetour = modAPI.RegQueryValueEx(plngHandle, pstrSousCléSource, 0, (plngTypeDonnées), pvarTampon, plngNombreOctets)
                Select Case plngRetour
'                   Si la sous-clé n'existe pas
                    Case apiERROR_FILE_NOT_FOUND
'                       Si la création de la sous-clé est exigée
                        If pbooCréerSiInexistante Then
                            pbooCléSousCléÀCréer = True
                        Else
'                           S'il existe une valeur par défaut pour la sous-clé
                            If Not IsMissing(pvarDéfaut) Then
                                pvarValeur = pvarDéfaut
                            Else
                                lngLectureValeurRegistre = plngRetour
                                plngRetour = modAPI.RegCloseKey(plngHandle)
                                Exit Function
                            End If
                        End If
'                   Si l'opération nécessite une deuxième lecture
                    Case apiERROR_MORE_DATA
                        pstrValeur = String(plngNombreOctets + 1, 0)
'                       Lire la valeur de la sous-clé
                        plngRetour = modAPI.RegQueryValueEx(plngHandle, pstrSousCléSource, 0, plngTypeDonnées, ByVal pstrValeur, plngNombreOctets)
                        If Not plngRetour = apiERROR_SUCCESS Then lngLectureValeurRegistre = plngRetour: plngRetour = modAPI.RegCloseKey(plngHandle): Exit Function
'                       En extraire l'entier
                        For plngIndiceOctet = plngNombreOctets To 1 Step -1
                            pstrValeurHexadécimale = pstrValeurHexadécimale & Format(Hex(Asc(Mid(pstrValeur, plngIndiceOctet, 1))), "@@")
                        Next plngIndiceOctet
                        pstrValeurHexadécimale = Replace(pstrValeurHexadécimale, " ", "0")
                        plngValeur = CDec("&H" & pstrValeurHexadécimale)
                        pvarValeur = plngValeur
                    Case Else
                        lngLectureValeurRegistre = plngRetour
                        plngRetour = modAPI.RegCloseKey(plngHandle)
                        Exit Function
                End Select
        
        End Select
        
    End If
    
'   Si la clé a dû être créée ou si la sous-clé n'existe pas
    If pbooCléSousCléÀCréer Then
'       Écrire la sous-clé et sa valeur par défaut
        plngRetour = lngÉcritureValeurRegistre(plngCléRacine, plngCléSource, pstrSousCléSource, plngTypeDonnées, pvarDéfaut)
        If Not plngRetour = apiERROR_SUCCESS Then
            lngLectureValeurRegistre = plngRetour
            plngRetour = modAPI.RegCloseKey(plngHandle)
            Exit Function
        Else
            pvarValeur = pvarDéfaut
        End If
    End If
    
'   Refermer la clé
    lngLectureValeurRegistre = modAPI.RegCloseKey(plngHandle)
    
    Exit Function
    
Err_lngLectureValeurRegistre:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
            
End Function


'Retourner un entier long précisant l'état d'une tentative d'ouverture de clé sur le registre
'
'   plngCléStandard:    index d'un nom de clé standard
'   plngCléSource:      chaîne de la clé sur laquelle porte l'ouverture
'
'   Valeurs retournées: apiERROR_SUCCESS:           la clé existe
'                       apiERROR_FILE_NOT_FOUND:    la clé n'existe pas dans le registre
'                       toute autre valeur:         l'ouverture a généré une erreur inattendue

Public Function lngVérificationExistenceCléRegistre(plngCléStandard As APIClésRacines, plngCléSource As String) As Long
Const pkstrNomProcédure As String = "lngVérificationExistenceCléRegistre"
Dim plngRetour As Long, plngHandle As Long
On Error GoTo Err_lngVérificationExistenceCléRegistre

'   Ouvrir la clé
    plngRetour = modAPI.RegOpenKeyEx(plngCléStandard, plngCléSource, 0, KEY_ALL_ACCESS, plngHandle)
    If Not plngRetour = apiERROR_SUCCESS Then lngVérificationExistenceCléRegistre = plngRetour: plngRetour = RegCloseKey(plngHandle): Exit Function
    
'   Refermer la clé
    lngVérificationExistenceCléRegistre = RegCloseKey(plngHandle)
    
    Exit Function
    
Err_lngVérificationExistenceCléRegistre:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Function


'Positionner une liste de DTPicker à la date du jour
'   pvarArgument(): tableau de type ParamArray constitué des DTPicker dont on souhaite le positionnement

Public Sub PositionnerDateJour(ParamArray pvarArgument() As Variant)
Const pkstrNomProcédure As String = "PositionnerDateJour"
Dim pdtpCible As DTPicker
Dim pintIndiceArgument As Integer
On Error GoTo Err_PositionnerDateJour

'   Pour chaque argument passé en référence
    For pintIndiceArgument = LBound(pvarArgument) To UBound(pvarArgument)
        Set pdtpCible = pvarArgument(pintIndiceArgument)
        pdtpCible.Value = Date
    Next pintIndiceArgument
    
    Exit Sub

Err_PositionnerDateJour:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Retourner une chaîne formatée constituée de la collection des erreurs ODBC

Public Function strChaîneErreursODBC() As String
Dim perrODBC As Error
Dim pstrODBCNuméro As String
Dim pintODBCLongueurNuméro As Integer, pintODBCLongueurNuméroMaximale As Integer, pintODBCIndice As Integer

    For Each perrODBC In Errors
        pintODBCLongueurNuméro = Len(Format(perrODBC.Number))
        If pintODBCLongueurNuméro > pintODBCLongueurNuméroMaximale Then pintODBCLongueurNuméroMaximale = pintODBCLongueurNuméro
    Next perrODBC
    
    For Each perrODBC In Errors
        pstrODBCNuméro = Format(perrODBC.Number)
        strChaîneErreursODBC = strChaîneErreursODBC & "Errors(" & Format(pintODBCIndice) & "):  " & Space(pintODBCLongueurNuméroMaximale - Len(pstrODBCNuméro)) & pstrODBCNuméro & "  *" & perrODBC.Description & vbCrLf
        pintODBCIndice = pintODBCIndice + 1
    Next perrODBC
    strChaîneErreursODBC = Left(strChaîneErreursODBC, Len(strChaîneErreursODBC) - (1 * Len(vbCrLf)))

End Function


'Retourner une valeur issue du formatage d'une chaîne en tenant compte du cas où la chaîne est de longueur nulle
'   pstrChaîne:                         chaîne sur laquelle porte l'opération
'   plngType:                           type de la valeur que l'on souhaite retourner
'   pvarValeurSiChaîneLongueurNulle:    valeur à retourner si la chaîne est de longueur nulle [facultatif]

Public Function varChaîneFormatée(pstrChaîne As String, plngType As VbVarType, Optional pvarValeurSiChaîneLongueurNulle As Variant) As Variant
Const pkstrNomProcédure As String = "varChaîneFormatée"
On Error GoTo Err_varChaîneFormatée

'   Si la chaîne est de longueur nulle et si la valeur à retourner est définie
    If pstrChaîne = "" And Not IsMissing(pvarValeurSiChaîneLongueurNulle) Then
        varChaîneFormatée = pvarValeurSiChaîneLongueurNulle
        Exit Function
    End If

'   Si la chaîne est de longueur nulle
    If pstrChaîne = "" Then
        Select Case plngType
'           Chaîne
            Case vbString:      varChaîneFormatée = ""
'           Date
            Case vbDate:        varChaîneFormatée = #12/31/9999#
'           Entier court
            Case vbInteger:     varChaîneFormatée = 0
'           Entier long
            Case vbLong:        varChaîneFormatée = 0
'           Nombre en simple précision
            Case vbSingle:      varChaîneFormatée = 0
'           Nombre en double précision
            Case vbDouble:      varChaîneFormatée = 0
'           Valeur monétaire
            Case vbCurrency:    varChaîneFormatée = 0
        End Select
'   Sinon
    Else
        Select Case plngType
'           Chaîne
            Case vbString:      varChaîneFormatée = pstrChaîne
'           Date
            Case vbDate:        varChaîneFormatée = CDate(pstrChaîne)
'           Entier court
            Case vbInteger:     varChaîneFormatée = CInt(Val(pstrChaîne))
'           Entier long
            Case vbLong:        varChaîneFormatée = CLng(Val(pstrChaîne))
'           Nombre en simple précision
            Case vbSingle:      varChaîneFormatée = CSng(Val(pstrChaîne))
'           Nombre en double précision
            Case vbDouble:      varChaîneFormatée = CDbl(Val(pstrChaîne))
'           Valeur monétaire
            Case vbCurrency:    varChaîneFormatée = CCur(Val(pstrChaîne))
        End Select
    End If
    
    Exit Function
    
Err_varChaîneFormatée:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Function



'Retourner une variable fonction du résultat de la recherche d'une valeur dans une collection de type chaîne, un contrôle ComboBox, un contrôle Mh3dCombo ou un contrôle ListView, ou d'une clé dans une collection Nodes
'
'   pobjCible:      objet sur lequel porte la recherche
'   pvarArgument(): tableau de type ParamArray constitué d'arguments implémentant le type et la nature de la recherche
'                   pour une collection de type chaîne:
'                       pvarArgument(0):    chaîne dont on souhaite la recherche
'                       pvarArgument(1):    index de l'élément à partir duquel effectuer la recherche
'                       pvarArgument(2):    numéro de la sous-chaîne sur laquelle doit s'effectuer la recherche dans le cas où l'élément de la collection est une chaîne formatée à l'aide du séparateur [facultatif: s'il est manquant, la recherche s'effectuera sur la totalité de l'élément]
'                   pour un contrôle ComboBox:
'                       pvarArgument(0):    booléen précisant si la recherche porte sur la propriété 'List' du contrôle [dans le cas contraire, la recherche portera sur la propriété 'ItemData' du contrôle]
'                       pvarArgument(1):    chaîne dont on souhaite la recherche
'                   pour un contrôle ListView:
'                       pvarArgument(0):    index du sous-élément où effectuer la recherche [s'il est égal à 0, la recherche se fera sur la propriété 'Text' des éléments du ListView]
'                       pvarArgument(1):    index de l'objet ListItem à partir duquel effectuer la recherche
'                       pvarArgument(2):    chaîne dont on souhaite la recherche
'                   pour une collection Nodes:
'                       pvarArgument(0):    clé dont on souhaite la recherche
'
'   Si la recherche est infructueuse, la valeur de la variable retournée est:
'       pour une collection de type chaîne: 0
'       pour un contrôle ComboBox:          comINDÉFINI
'       pour un contrôle ListView:          Nothing
'       pour une collection Nodes:          Nothing
'   Dans le cas contraire, la valeur de la variable retournée est:
'       pour une collection de type chaîne: l'index de l'élément contenant la valeur recherchée
'       pour un contrôle ComboBox:          l'index de la ligne contenant la valeur recherchée
'       pour un contrôle Mh3dCombo:         l'index de la ligne contenant la valeur recherchée
'       pour un contrôle ListView:          l'objet ListItem contenant la valeur recherchée
'       pour une collection Nodes:          le noeud contenant la clé recherchée

Public Function varRecherche(pobjCible As Object, ParamArray pvarArgument() As Variant) As Variant
Const pkstrNomProcédure As String = "varRecherche"
Dim plvwSource As ListView, pitmSource As ListItem, pcolSource As Collection, pnodSources As Nodes, pnodSource As Node, pcbosource As ComboBox
Dim pbooErreurPossible As Boolean, pintIndex As Integer, pstrCollection As String, pintPositionSéparateur As Integer, pintNombreRéductions As Integer
On Error GoTo Err_varRecherche

'   S'il s'agit d'une collection
    If TypeOf pobjCible Is Collection Then
'       (pvarArgument: (0)=chaîne à rechercher, (1)=index du membre à partir duquel effectuer la recherche, (2)=sous-chaîne ou effectuer la recherche)
        Set pcolSource = pobjCible
        With pcolSource
            varRecherche = 0
            For pintIndex = pvarArgument(1) To pcolSource.Count
                pstrCollection = pcolSource(pintIndex)
'               Si la collection est formatée
                If UBound(pvarArgument) = 2 Then
                    pintNombreRéductions = 0
'                   Localiser la zone de recherche
                    Do
                        pintPositionSéparateur = InStr(pstrCollection, gstrSÉPARATEUR_1)
                        pstrCollection = Right(pstrCollection, Len(pstrCollection) - pintPositionSéparateur)
                        pintNombreRéductions = pintNombreRéductions + 1
                    Loop Until pintNombreRéductions = pvarArgument(2)
                End If
                If InStr(pstrCollection, pvarArgument(0)) = 1 Then varRecherche = pintIndex: Exit For
            Next pintIndex
        End With
'   S'il s'agit d'un ComboBox
    ElseIf TypeOf pobjCible Is ComboBox Then
'       (pvarArgument: (0)=booléen précisant la portée de la recherche, (1)=chaîne à rechercher)
        varRecherche = comINDÉFINI
        If Not IsNull(pvarArgument(1)) Then
            Set pcbosource = pobjCible
            With pcbosource
                For pintIndex = 0 To pcbosource.ListCount - 1
'                   Si la recherche porte sur la propriété 'List'
                    If pvarArgument(0) Then
                        If pvarArgument(1) = .List(pintIndex) Then varRecherche = pintIndex: Exit For
'                   Si la recherche porte sur la propriété 'ItemData'
                    Else
                        If pvarArgument(1) = .ItemData(pintIndex) Then varRecherche = pintIndex: Exit For
                    End If
                Next pintIndex
            End With
        End If
'   S'il s'agit d'un ListView
    ElseIf TypeOf pobjCible Is ListView Then
'       (pvarArgument: (0)=Text/SubItem où effectuer la recherche, (1)=position de début de recherche, (2)=chaîne à rechercher)
        Set plvwSource = pobjCible
        If plvwSource.ListItems.Count = 0 Then
            Set pitmSource = Nothing
        Else
'           S'il s'agit d'effectuer la recherche sur la propriété 'Text' des éléments du ListView
            If pvarArgument(0) = 0 Then
                Set pitmSource = plvwSource.FindItem(pvarArgument(2), lvwText, pvarArgument(1))
            Else
                Do
                    Set pitmSource = plvwSource.FindItem(pvarArgument(2), lvwSubItem, pvarArgument(1))
'                   Si aucune chaîne n'est trouvée, sortir de la boucle
                    If pitmSource Is Nothing Then Exit Do
'                   S'il s'agit du dernier élément du ListView et que la chaîne trouvée n'appartient pas au SubItem référencé
                    If pitmSource.Index = plvwSource.ListItems.Count And Not pitmSource.SubItems(pvarArgument(0)) = pvarArgument(2) Then
                        Set pitmSource = Nothing
                        Exit Do
                    End If
                    pvarArgument(1) = pitmSource.Index + 1
'               Boucler tant que la chaîne est trouvée mais qu'elle n'appartient pas au SubItem référencé
                Loop Until pitmSource.SubItems(pvarArgument(0)) = pvarArgument(2)
            End If
        End If
'       Retourner le résultat de la recherche
        Set varRecherche = pitmSource
'   S'il s'agit d'un TreeView
    ElseIf TypeOf pobjCible Is Nodes Then
'       (pvarArgument: (0)=clé à rechercher)
        Set pnodSources = pobjCible
        pbooErreurPossible = True
'       Référencer le noeud recherché
'           Génère l'erreur 35601 (ccElemNotFound) si la clé est inexistante
        Set pnodSource = pnodSources(pvarArgument(0))
        pbooErreurPossible = False
'       Retourner le résultat de la recherche
        Set varRecherche = pnodSource
    End If
    
    Exit Function
    
Err_varRecherche:
Dim pintTraitement As MessagesActions
    If pbooErreurPossible And Err.Number = ccElemNotFound Then Resume Next
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Function


'Retourner une valeur après prise en compte de la valeur 'Null' et formatage éventuel
'   pvarValeur:         valeur sur laquelle porte l'opération
'   plngType:           type de la valeur que l'on souhaite retourner
'   pvarValeurSiNull:   valeur à retourner si pvarValeur est Null [facultatif]
'   pvarFormat:         format souhaité [facultatif]

Public Function varValeurFormatéeSiNull(pvarValeur As Variant, plngType As VbVarType, Optional pvarValeurSiNull As Variant, Optional pvarFormat As Variant) As Variant
Const pkstrNomProcédure As String = "varValeurFormatéeSiNull"
On Error GoTo Err_varValeurFormatéeSiNull

'   S'il s'agit d'une valeur nulle et si la valeur à retourner est définie
    If IsNull(pvarValeur) And Not IsMissing(pvarValeurSiNull) Then
        varValeurFormatéeSiNull = pvarValeurSiNull
        Exit Function
    End If

'   S'il s'agit d'une valeur nulle
    If IsNull(pvarValeur) Then
        Select Case plngType
'           Chaîne
            Case vbString:                                          varValeurFormatéeSiNull = ""
'           Date
            Case vbDate:                                            varValeurFormatéeSiNull = #12/31/9999#
'           Entier court ou long, nombre en simple ou double précision, ou valeur monétaire
            Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency: varValeurFormatéeSiNull = 0
        End Select
'   Sinon
    Else
        varValeurFormatéeSiNull = pvarValeur
    End If
    
'   Formater la valeur si le formatage est demandé
    If Not IsMissing(pvarFormat) Then varValeurFormatéeSiNull = Format(varValeurFormatéeSiNull, pvarFormat)

    Exit Function
    
Err_varValeurFormatéeSiNull:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Function


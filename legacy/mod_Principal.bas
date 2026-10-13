Attribute VB_Name = "mod_Principal"
Option Explicit

Private Const mstrNOM_MODULE As String = "mod_Principal"

Public CheminBase As String
Public Db As Database
Public DbAccessPath As String 'chemin de la base utilisée pour les mailings
Public DbAccess As Database   'base utilisée pour les mailings
Public DbAccessResPath As String  'chemin de la base utilisée pour le résultat de la sélection (mailings)
Public DbAccessRes As Database    'base utilisée pour le résultat de la sélection des enregistrements (mailings)
Public FichierDonnées As String   'fichier de données utilisé pour les courriers
Public CheminWord As String 'chemin des modèles word

Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'ajouté le 13/01/05 par Caroline pour XP
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public RepertoireTemporaire As String


Public Sub RecupRepertoireTemporaire()
'ajouté le 13/01/05 par Caroline pour XP

  Dim Valeur As String
  Dim longueur As Long
  
  Valeur = String$(100, " ")
  longueur = GetTempPath(Len(Valeur), Valeur)
  RepertoireTemporaire = UCase(Left$(Valeur, longueur))

End Sub


Public Function IifNull(ChDate As Variant) As String

  'IifNull = IIf(Not IsNull(ChDate), """" + CStr(ChDate) + """", """NULL""")
  If Not IsNull(ChDate) Then
    IifNull = """" + CStr(ChDate) + """"
  Else
    IifNull = "NULL"
  End If
  
End Function

Public Sub ListeDossiers(Dat_ctrl As Data, Boite As ComboBox)
'Place la liste des services dans la combo
    With Dat_ctrl
      .RecordSource = "SELECT DOSS_NUM FROM DOSS ORDER BY DOSS_NUM DESC"
      .Refresh
      Boite.Clear
      While Not Dat_ctrl.Recordset.EOF
        Boite.AddItem (.Recordset.Fields("DOSS_NUM"))
        .Recordset.MoveNext
      Wend
    End With
End Sub


Public Sub ListeDemandesPourCeDossier(Dat_ctrl As Data, Boite As ComboBox, Dossier As String)
'Place la liste des demandes relatives au dossier n°XX dans la combo
    With Dat_ctrl
        .RecordSource = "SELECT DEMA_NUM,DEMA_DATE FROM DEMA WHERE DOSS_NUM = " + Iifch(Dossier)
        .Refresh
        Boite.Clear
        While Not .Recordset.EOF
            Boite.AddItem (.Recordset.Fields("DEMA_NUM"))
            .Recordset.MoveNext
        Wend
    End With
End Sub
Public Sub OuvrirFichier(NomFichier As String)
  
  Dim hProcess As Long
  Dim hWndDesk As Long
  Dim ProcessId As Long
  
  hWndDesk = GetDesktopWindow()
  ProcessId = ShellExecute(hWndDesk, "Open", NomFichier, "", 0&, 1)

End Sub

Sub OuvrirWord()

Dim WdApp As Object

On Error GoTo PasOuvert
    
  Set WdApp = GetObject(, "word.application")
  WdApp.Visible = True
  Set WdApp = Nothing
  Exit Sub
  
PasOuvert:
  Set WdApp = CreateObject("word.application")
  WdApp.Visible = True
  Set WdApp = Nothing
  
End Sub

Public Sub Tri(Dat_ctrl As Object, dbgrid As Object, colonne As String)

'Permet le tri en fonction de la colonne sélectionnée ( tri croissant puis décroissant )
'dbgrid est le DBGRIB sur lequel sera effectué le tri
'dat_ctrl est le CONTROLE_DATA source du dbgrid
'colonne est le nom de la colonne sur laquelle va être effectué le tri
'Les recordsources spécifiés au départ doivent toujours être sous forme de requêtes SQL

  Dim classem As String
  Dim c As Integer
  Dim t As Integer
  Dim j As Integer
  
  
    t = InStr(Dat_ctrl.RecordSource, " ORDER BY")
    If t = 0 Then t = InStr(Dat_ctrl.RecordSource, " order by")
    
    If t = 0 Then
        'il n'y a pas de ORDER BY dans la requête
        Dat_ctrl.RecordSource = Dat_ctrl.RecordSource + " ORDER BY " + colonne + " ASC"
    
    Else
        'il existe déja un ORDER BY dans la requête
        'on regarde sur quelle colonne est effectuée le tri
        c = InStr(Dat_ctrl.RecordSource, "ORDER BY " + colonne + " ")
        
        If c <> 0 Then
            'le tri est effectué sur la même colonne
            j = InStr(Dat_ctrl.RecordSource, "ASC")
            
            If j <> 0 Then
                 Dat_ctrl.RecordSource = Left(Dat_ctrl.RecordSource, j - 1) + "DESC"
            Else
                 j = InStr(Dat_ctrl.RecordSource, "DESC")
                 Dat_ctrl.RecordSource = Left(Dat_ctrl.RecordSource, j - 1) + "ASC"
            End If
            
        Else
            'le tri est effectué sur une autre colonne
            Dat_ctrl.RecordSource = Left(Dat_ctrl.RecordSource, t) + "ORDER BY " + colonne + " ASC"
        End If
        
    End If
        
    'raffraichissement du contrôle data
    Dat_ctrl.Refresh

End Sub

Public Sub CentrageHorizontal(Fenêtre As Object)
  Fenêtre.Left = (Screen.Width - Fenêtre.Width) \ 2
End Sub
'Procédure Main : programme principal
Public Sub Main()
Const pkstrNomProcédure As String = "Main"
Dim WdApp As Object
On Error GoTo Err_Main

    Screen.MousePointer = vbHourglass

'   Initialiser l'application
    modAdoptions.Initialiser
    
    Screen.MousePointer = vbDefault

    CheminBase = modDAO.strNomBaseAvecChemin()
    
'   Lister les demandes à réactualiser
    frmDemandesÀRéactualiser.Hide
    frmDemandesÀRéactualiser.GérerAffichageFenêtre
    
  'Ouverture de Word
  OuvrirWord
  
  'Ouverture de la base de données
  Set Db = Workspaces(0).OpenDatabase(CheminBase)
  
  DbAccessPath = "c:\program files\" + App.EXEName + "\"
  DbAccessResPath = "c:\program files\" + App.EXEName + "\cg22.mdb"
  FichierDonnées = "c:\program files\" + App.EXEName + "\cg22.doc"
  
  'Récupération du chemin des modèles Word
  Menu.Dat_CheminWord.DatabaseName = CheminBase
  Menu.Dat_CheminWord.RecordSource = "SELECT * FROM PARAM WHERE PARAM_Code=1"
  Menu.Dat_CheminWord.Refresh
  If Menu.Dat_CheminWord.Recordset.RecordCount = 0 Then
    MsgBox "Chemin des modèles Word non spécifié", vbOKOnly, vbInformation
  Else
    CheminWord = Menu.Dat_CheminWord.Recordset.Fields(2)
  End If
  
  'Chargement de la fenêtre principale
  Load Menu
  Menu.Show vbModeless
    
    Exit Sub
    
Err_Main:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Sub


Public Function Iifch(textbox As String) As String

    Iifch = IIf(textbox <> "", """" + textbox + """", "NULL")
    
End Function
Public Function Iifent(textbox As String) As String

    Iifent = IIf(textbox <> "", CStr(textbox), "NULL")
    
End Function

Function PresentDansListBox(Boite As Object, Chaine As String)
'Renvoie l'index d'une chaine dans une liste, -1 si absente

Dim Present As Boolean
Dim Compteur As Integer

  Present = False

  If Boite.ListCount <> 0 Then
    Compteur = 0
     While ((Compteur < Boite.ListCount) And (Not Present))
      If InStr(Boite.List(Compteur), Chaine) <> 0 Then Present = True
      Compteur = Compteur + 1
    Wend
  End If

  If Present Then
    PresentDansListBox = Compteur - 1
  Else
    PresentDansListBox = -1
  End If
  
End Function


Function CodeEntreCotes(Chaine As String)
Dim code As String
  If Chaine <> "" Then
      code = Mid(Chaine, InStr(Chaine, "<") + 1)
      code = Mid(code, 1, Len(code) - 1)
      CodeEntreCotes = code
  End If
End Function

Function NouveauCode(Table As String, Optional Champ As String = "")
'Renvoie automatiquement un nouveau code numérique pour la table en entrée
'Si le champ n'est pas spécifié, on prend Table + _CODE
'On démarre à 1

Dim Chp As String

  If Champ = "" Then Chp = Table + "_CODE" Else Chp = Champ
  Menu.Dat_CodeGenerateur.RecordSource = "SELECT MAX(" + Chp + ") FROM " + Table
  Menu.Dat_CodeGenerateur.Refresh
  If ((Menu.Dat_CodeGenerateur.Recordset.RecordCount = 0) Or (IsNull(Menu.Dat_CodeGenerateur.Recordset.Fields(0)))) Then
    NouveauCode = 1
  Else
    NouveauCode = Menu.Dat_CodeGenerateur.Recordset.Fields(0) + 1
  End If
End Function

Function NouveauCodeNouveauDossier()

Dim annee As String
Dim numero As String
Dim NumMax As String

  Menu.Dat_CodeGenerateur.RecordSource = "SELECT MAX(DOSS_NUM) FROM DOSS"
  Menu.Dat_CodeGenerateur.Refresh
  If ((Menu.Dat_CodeGenerateur.Recordset.RecordCount = 0) Or (IsNull(Menu.Dat_CodeGenerateur.Recordset.Fields(0)))) Then
    NouveauCodeNouveauDossier = CStr(Year(Date)) * 1000 + CStr(1)
  Else
    NumMax = Menu.Dat_CodeGenerateur.Recordset.Fields(0)
    annee = Left(NumMax, 4)
    numero = Mid(NumMax, 5, Len(NumMax))
    If annee >= Year(Date) Then
      numero = numero + 1
      NouveauCodeNouveauDossier = CStr(NumMax) + 1
    Else
      numero = 1
      annee = Year(Date)
      NouveauCodeNouveauDossier = CStr(annee) * 1000 + CStr(numero)
    End If
  End If
End Function

Public Sub ListeFratrie(Boite As ComboBox)
'Place la liste des code de Fratrie dans la combo
    Boite.AddItem ("Demandeur <DE>")
    Boite.AddItem ("Conjoint  <CO>")
    Boite.AddItem ("Enfant au foyer <EF>")
    Boite.AddItem ("Enfant hors foyer <EH>")
    Boite.AddItem ("Autre personne <AP>")
End Sub

Public Sub ListeAssistantesSociales(Dat_ctrl As Data, Boite As ComboBox)
'Place la liste des assistantes sociales dans la combo
    With Dat_ctrl
        .RecordSource = "SELECT TRSO_NUM,TRSO_NOM FROM TRSO WHERE QUAL_CODE IN (" + _
            "SELECT QUAL_CODE FROM QUAL WHERE QUAL_LIB = ""ASSISTANTE SOCIALE"")"
        .Refresh
        Boite.Clear
        Boite.AddItem ("")
        While Not .Recordset.EOF
            Boite.AddItem (.Recordset.Fields("TRSO_NOM") + " <" + CStr(.Recordset.Fields("TRSO_NUM")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub


Public Sub ListePsychologues(Dat_ctrl As Data, Boite As ComboBox)
'Place la liste des psychologues dans la combo
    With Dat_ctrl
        .RecordSource = "SELECT TRSO_NUM,TRSO_NOM FROM TRSO WHERE QUAL_CODE IN (" + _
            "SELECT QUAL_CODE FROM QUAL WHERE QUAL_LIB = ""PSYCHOLOGUE"")"
        .Refresh
        Boite.Clear
        Boite.AddItem ("")
        While Not .Recordset.EOF
            Boite.AddItem (.Recordset.Fields("TRSO_NOM") + " <" + CStr(.Recordset.Fields("TRSO_NUM")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub

Public Sub ListeRecoursContencieux(Dat_ctrl As Data, Boite As ComboBox)
'Place la liste des tribunaux auprès desquels peuvent être effectué un recours contencieux dans la combo
    With Dat_ctrl
        .RecordSource = "SELECT * FROM TYRE"
        .Refresh
        Boite.Clear
        While Not .Recordset.EOF
            Boite.AddItem (.Recordset.Fields("TYRE_NOM") + " <" + CStr(.Recordset.Fields("TYRE_NUM")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub
Public Sub ListeAccordRefus(Boite As ComboBox)
'Place la liste des codes de décision dans la combo (A=Accord,R=Refus)
    Boite.AddItem ("")
    Boite.AddItem ("Accord <A>")
    Boite.AddItem ("Refus <R>")
End Sub

Public Sub ExportExcel(DataSource As String)
  
  Dim Db As Database
  Dim Rs As Recordset
  Dim resultat As Long
  Dim curseur As Integer
  Dim fic As String
  
  Screen.MousePointer = vbHourglass
  
  Set Db = Workspaces(0).OpenDatabase(CheminBase)
  Set Rs = Db.OpenRecordset(DataSource)
  
  If Rs.EOF Then
    Screen.MousePointer = vbDefault
    Exit Sub
  Else
    Rs.MoveFirst
  End If
  
  'ajouté le 13/01/05 par Caroline pour XP
  Call RecupRepertoireTemporaire
  fic = RepertoireTemporaire & "Adop.xls"
  Open fic For Output As #1
  'Open "c:\windows\temp\Adop.xls" For Output As #1
  
  'enregistrement des noms de champs
  
  For curseur = 0 To Rs.Fields.Count - 1
      Print #1, Rs.Fields(curseur).Name; Chr(9);
  Next curseur
  
  Print #1,
  
  'impression des enregistrements
  
  Rs.MoveFirst
  
  Do While Not Rs.EOF
  
    For curseur = 0 To Rs.Fields.Count - 1
      'Print #1, """" + CStr(Rs.Fields(curseur).Value) + """"; Chr(9)
      If Not IsNull(Rs.Fields(curseur).Value) Then
        Print #1, """" + CStr(Rs.Fields(curseur).Value) + """"; Chr(9);
      Else
        Print #1, ""; Chr(9);
      End If
    Next curseur
    
    
    Print #1,
    
    Rs.MoveNext
    
  Loop
  
  Close #1
  Rs.Close
  Db.Close
  
  Screen.MousePointer = vbDefault
  
  'lancement d'excel
  'modifié le 13/01/05 par Caroline
  'Call OuvrirFichier("c:\windows\temp\Adop.xls")
  Call OuvrirFichier(fic)
  
End Sub

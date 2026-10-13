VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDécisions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const mstrNOM_MODULE_CLASSE As String = "clsDécisions"

Private Enum Erreurs    'Erreurs Visual Basic
    ERREUR_00005 = 5    'Argument ou appel de procédure incorrect
    ERREUR_00457 = 457  'Cette clé est déjà associée à un élément de cette collection
End Enum

Private mcolDécisions As Collection


Public Function Ajout(plngErreur As Long, pdtmCommissionInitiale As Date, plngProvenance As Long, plngDécision As Long, pvarCommission As Variant, pvarSignature As Variant, Optional pvarDépôt As Variant = Null, Optional pvarClé As Variant, Optional pvarAvant As Variant) As clsDécision
Const pkstrNomProcédure As String = "Ajout"
Dim pclsNouvelÉlément As clsDécision
On Error GoTo Err_Ajout
    
    Set pclsNouvelÉlément = New clsDécision
    
    With pclsNouvelÉlément
        .dtmCommissionInitiale = pdtmCommissionInitiale
        .lngProvenance = plngProvenance
        .lngDécision = plngDécision
        .varCommission = pvarCommission
        .varSignature = pvarSignature
        .varDépôt = pvarDépôt
        If Not IsMissing(pvarClé) Then .strClé = pvarClé
    End With

    If IsMissing(pvarClé) Then
'       Génère l'erreur 457 si l'élément existe déjà
        If IsMissing(pvarAvant) Then mcolDécisions.Add pclsNouvelÉlément Else mcolDécisions.Add pclsNouvelÉlément, Before:=pvarAvant
    Else
'       Génère l'erreur 457 si l'élément existe déjà
        If IsMissing(pvarAvant) Then mcolDécisions.Add pclsNouvelÉlément, Key:=pvarClé Else mcolDécisions.Add pclsNouvelÉlément, Key:=pvarClé, Before:=pvarAvant
    End If
    
    Set Ajout = pclsNouvelÉlément
    
    Set pclsNouvelÉlément = Nothing

    Exit Function
    
Err_Ajout:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:         pintTraitement = mesINTERROMPRE
        Case Erreurs.ERREUR_00457:  plngErreur = Err.Number: Resume Next
        Case Else:                  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE_CLASSE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Function


Public Sub Effacer()
Const pkstrNomProcédure As String = "Effacer"
Dim pintIndex As Integer
On Error GoTo Err_Effacer
    
'   Effacer la collection
    With mcolDécisions
        For pintIndex = .Count To 1 Step -1
            .Remove pintIndex
        Next pintIndex
    End With

    Exit Sub
    
Err_Effacer:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE_CLASSE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Sub


Public Property Get Élément(pvarIndex As Variant) As clsDécision
Attribute Élément.VB_UserMemId = 0
Const pkstrNomProcédure As String = "Élément"
On Error GoTo Err_Élément
  
'   Génère l'erreur 5 si l'élément n'existe pas
    Set Élément = mcolDécisions(pvarIndex)
    
    Exit Property
    
Err_Élément:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:         pintTraitement = mesINTERROMPRE
        Case Erreurs.ERREUR_00005:  Resume Next
        Case Else:                  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE_CLASSE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Property


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const pkstrNomProcédure As String = "NewEnum"
On Error GoTo Err_NewEnum
    
    Set NewEnum = mcolDécisions.[_NewEnum]

    Exit Property
    
Err_NewEnum:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE_CLASSE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Property


Public Property Get Nombre() As Long
Const pkstrNomProcédure As String = "Nombre"
On Error GoTo Err_Nombre
    
    Nombre = mcolDécisions.Count

    Exit Property
    
Err_Nombre:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE_CLASSE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Property


Public Sub Supprimer(pvarIndex As Variant)
Const pkstrNomProcédure As String = "Supprimer"
On Error GoTo Err_Supprimer

'   Supprimer l'élément de la collection désigné par sa clé ou son index
    mcolDécisions.Remove pvarIndex

    Exit Sub
    
Err_Supprimer:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError:         pintTraitement = mesINTERROMPRE
        Case Erreurs.ERREUR_00005:  Resume Next
        Case Else:                  pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE_CLASSE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Sub


Private Sub Class_Initialize()
Const pkstrNomProcédure As String = "Class_Initialize"
On Error GoTo Err_Class_Initialize
    
    Set mcolDécisions = New Collection

    Exit Sub
    
Err_Class_Initialize:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE_CLASSE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Sub


Private Sub Class_Terminate()
Const pkstrNomProcédure As String = "Class_Terminate"
On Error GoTo Err_Class_Terminate
    
    Set mcolDécisions = Nothing

    Exit Sub
    
Err_Class_Terminate:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE_CLASSE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Sub


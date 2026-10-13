VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecoursContentieux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const mstrNOM_MODULE_CLASSE As String = "clsRecoursContentieux"

Public lngTribunalAdministratif As Long
Public lngCourAdministrativeAppel As Long
Public lngConseilÉtat As Long
Public lngPlusHautRecours As Long
Public clsRecoursContentieux As clsCodesLibellés

Private Const mstrTRIBUNAL_ADMINISTRATIF As String = "Tribunal Administratif", mstrCOUR_ADMINISTRATIVE_APPEL As String = "Cour Administrative d'Appel", mstrCONSEIL_ÉTAT As String = "Conseil d'État"


'Retourner un booléen précisant si le paramétrage des recours contentieux est complet

Public Function booParamétrageIncomplet() As Boolean
Const pkstrNomProcédure As String = "booParamétrageIncomplet"
On Error GoTo Err_booParamétrageIncomplet
    
    booParamétrageIncomplet = (lngTribunalAdministratif = comINDÉFINI Or lngCourAdministrativeAppel = comINDÉFINI Or lngConseilÉtat = comINDÉFINI)

    Exit Function
    
Err_booParamétrageIncomplet:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE_CLASSE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Function


Public Property Get strConseilÉtat() As String
Const pkstrNomProcédure As String = "strConseilÉtat"
On Error GoTo Err_strConseilÉtat
  
    strConseilÉtat = mstrCONSEIL_ÉTAT
    
    Exit Property
    
Err_strConseilÉtat:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE_CLASSE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Property



Public Property Get strCourAdministrativeAppel() As String
Const pkstrNomProcédure As String = "strCourAdministrativeAppel"
On Error GoTo Err_strCourAdministrativeAppel
  
    strCourAdministrativeAppel = mstrCOUR_ADMINISTRATIVE_APPEL
    
    Exit Property
    
Err_strCourAdministrativeAppel:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE_CLASSE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Property



Public Property Get strTribunalAdministratif() As String
Const pkstrNomProcédure As String = "strTribunalAdministratif"
On Error GoTo Err_strTribunalAdministratif
  
    strTribunalAdministratif = mstrTRIBUNAL_ADMINISTRATIF
    
    Exit Property
    
Err_strTribunalAdministratif:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, mstrNOM_MODULE_CLASSE, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDécision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const mstrNOM_MODULE_CLASSE As String = "clsDécision"

Public strClé As String
Public dtmCommissionInitiale As Date
Public lngProvenance As Long
Public lngDécision As Long
Public varCommission As Variant
Public varSignature As Variant
Public varDépôt As Variant


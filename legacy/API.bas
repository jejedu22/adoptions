Attribute VB_Name = "modAPI"
Option Explicit

Private Const mstrNOM_MODULE As String = "modAPI"

Public Enum APIErreurs
    apiERROR_SUCCESS = 0&
    apiERROR_FILE_NOT_FOUND = 2&
    apiERROR_MORE_DATA = 234
End Enum
Public Enum APIClésRacines
    apiHKEY_CURRENT_USER = &H80000001
    apiHKEY_LOCAL_MACHINE
End Enum
Public Enum APITypeDonnées
    apiREG_SZ = 1
    apiREG_DWORD = 4
End Enum

Public Type POINTAPI
    apiX As Long
    apiY As Long
End Type

Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_NOTIFY = &H10
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Public Const REG_OPTION_NON_VOLATILE = 0

Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2

' Récupère le nom de l'ordinateur
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
' Renvoie la position courante du curseur
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' Récupère le nom de l'utilisateur
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
' Recherche la première fenêtre de plus haut niveau correspondant aux conditions spécifiées dans la liste des fenêtres
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
' Récupère la chaîne pour l'entrée spécifiée dans un fichier d'initialisation
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
' Détermine si une fenêtre est l'enfant ou la descendante d'une autre
Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long
' Détermine si une fenêtre est visible
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
' Poste un message dans la file de messages de la fenêtre spécifiée
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' Referme une clé dans la base des registres système
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
' Crée une nouvelle clé sous la clé spécifiée; si la clé existe déjà, ouvre cette clé
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
' Ouvre une clé existante
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
' Récupère la valeur d'une sous-clé
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
' Définit la valeur d'une sous-clé
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
' Définit la position du curseur
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
' Renvoie le handle de la fenêtre contenant le point spécifié; ignore les fenêtres désactivées, cachées ou transparentes
Declare Function WindowFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
' Définit une chaîne pour l'entrée spécifiée dans un fichier d'initialisation
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


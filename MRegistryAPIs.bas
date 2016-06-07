Attribute VB_Name = "MRegistryAPIs"
Option Explicit
   
'Key APIs

Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, _
ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
Long) As Long

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'Value APIs

Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, _
ByVal lpData As String, lpcbData As Long) As Long

Declare Function RegQueryValueExLongRef Lib "advapi32.dll" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, _
lpData As Long, lpcbData As Long) As Long

Declare Function RegQueryValueExLongVal Lib "advapi32.dll" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, _
ByVal lpData As Long, lpcbData As Long) As Long

Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, _
ByVal lpValue As String, ByVal cbData As Long) As Long

Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, _
lpValue As Long, ByVal cbData As Long) As Long

'Delete APIs

Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String) As Long

Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
(ByVal hKey As Long, ByVal lpValueName As String) As Long

'Constants

Public Const REG_SZ As Long = 1
Public Const REG_DWORD As Long = 4
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_ARENA_TRASHED = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259

Public Const KEY_ALL_ACCESS = &H3F

Public Const REG_OPTION_NON_VOLATILE = 0

Public Sub Reg_CreateRegistryKey(pl_TopLevelKey As Long, ps_Key As String)
' e.g. CreateRegistryKey "NewKey", HKEY_CURRENT_USER
' CreateRegistryKey "NewKey\NewSubKey", HKEY_LOCAL_MACHINE

   Dim ll_KeyHandle As Long
   Dim ll_FunctionPerformed As Long
   
   'Create Key...
   'This key is not volatile; the information is stored in a file and is preserved when the system is restarted.
   If RegCreateKeyEx(pl_TopLevelKey, ps_Key, 0&, vbNullString, _
      REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, ll_KeyHandle, ll_FunctionPerformed) _
      = ERROR_NONE Then
           'Find out if it really is a new key or not...
           Select Case ll_FunctionPerformed
               Case REG_CREATED_NEW_KEY
                   MsgBox "Created New Key"
               Case REG_OPENED_EXISTING_KEY
                   MsgBox "Opened Existing Key"
           End Select
           
           RegCloseKey (ll_KeyHandle)
       
   Else
       MsgBox "Couldn't Create Key!"
   End If
   
End Sub

Public Sub Reg_DeleteRegistryKey(pl_TopLevelKey As Long, ps_Key As String)
' e.g. Reg_DeleteRegistryKey "NewKey", HKEY_CURRENT_USER
' Reg_DeleteRegistryKey "NewKey\NewSubKey", HKEY_LOCAL_MACHINE
   
   'Delete Key...
   If RegDeleteKey(pl_TopLevelKey, ps_Key) <> ERROR_NONE Then
       MsgBox "Couldn't Delete Key!"
   Else
       MsgBox "Deleted Key."
   End If
   
End Sub

Public Sub Reg_SetRegistryValue(pl_TopLevelKey As Long, ps_Key As String, ps_ValueName As String, _
   pv_Value As Variant, lValueType As Long)
   ' e.g. SetRegistryValue HKEY_CURRENT_USER, "Key\SubKey", "Name", "Lewis", REG_SZ
   ' SetRegistryValue HKEY_LOCAL_MACHINE, "Key\SubKey", "Age", 27, REG_DWORD
   
   Dim ll_KeyHandle As Long
   Dim ll_ReturnValue As Long
   Dim ll_Value As Long
   Dim ls_Value As String
   
   'Open Key
   If RegOpenKeyEx(pl_TopLevelKey, ps_Key, 0&, KEY_ALL_ACCESS, ll_KeyHandle) = ERROR_NONE Then
       
       'Set value (need to do differently for strings and longs)
       Select Case lValueType
          Case REG_SZ
              ls_Value = pv_Value & vbNullChar
              ll_ReturnValue = RegSetValueExString(ll_KeyHandle, _
                   ps_ValueName, 0&, lValueType, ls_Value, Len(ls_Value))
          Case REG_DWORD
              ll_Value = pv_Value
              ll_ReturnValue = RegSetValueExLong(ll_KeyHandle, _
                   ps_ValueName, 0&, lValueType, ll_Value, 4)
       End Select
       
       If ll_ReturnValue = ERROR_NONE Then
       Else
           MsgBox "Couldn't Change Value!"
       End If
       
       RegCloseKey (ll_KeyHandle)
   Else
       MsgBox "Couldn't Open Key!"
   End If
   
End Sub

Public Sub Reg_DeleteRegistryValue(pl_TopLevelKey As Long, ps_Key As String, ps_ValueName As String)
   ' e.g. Reg_DeleteRegistryValue HKEY_CURRENT_USER, "Key\SubKey", "Name"
   ' Reg_DeleteRegistryValue HKEY_LOCAL_MACHINE, "Key\SubKey", "Age"
   
   Dim ll_KeyHandle As Long
   Dim ll_ReturnValue As Long
   
   'Delete Key
   If RegOpenKeyEx(pl_TopLevelKey, ps_Key, 0&, KEY_ALL_ACCESS, ll_KeyHandle) = ERROR_NONE Then
       If RegDeleteValue(ll_KeyHandle, ps_ValueName) <> ERROR_NONE Then
           MsgBox "Couldn't Delete Value!"
       Else
           MsgBox "Deleted Value."
       End If
       RegCloseKey (ll_KeyHandle)
   Else
       MsgBox "Couldn't Open Key!"
   End If
   
End Sub

Public Function Reg_QueryRegistryValue(pl_TopLevelKey As Long, ps_Key As String, ps_ValueName As String) _
   As Variant
   ' x = QueryValue ("Key\SubKey", "Name")
   
   Dim ll_KeyHandle As Long
   Dim lv_Value As Variant
   Dim ll_Type As Long
   Dim ll_BufferSize As Long
   Dim ll_Value As Long
   Dim ls_Value As String

   
   'Open Key
   If RegOpenKeyEx(pl_TopLevelKey, ps_Key, 0&, KEY_ALL_ACCESS, ll_KeyHandle) = ERROR_NONE Then
        
        'Find length and type of data
        If RegQueryValueExLongVal(ll_KeyHandle, ps_ValueName, 0&, ll_Type, 0&, ll_BufferSize) _
           = ERROR_NONE Then

           'Set value (need to do differently for strings and longs)
           Select Case ll_Type
               Case REG_SZ
               
                   ls_Value = Space(ll_BufferSize)
                   If RegQueryValueExString(ll_KeyHandle, ps_ValueName, 0&, ll_Type, _
                       ls_Value, ll_BufferSize) = ERROR_NONE Then
                       Reg_QueryRegistryValue = Left(ls_Value, ll_BufferSize - 1)
                   Else
                       MsgBox "Couldn't get String Value!"
                   End If
       
               Case REG_DWORD
                    
                   If RegQueryValueExLongRef(ll_KeyHandle, ps_ValueName, 0&, ll_Type, _
                      ll_Value, ll_BufferSize) = ERROR_NONE Then
                      Reg_QueryRegistryValue = ll_Value
                   Else
                      MsgBox "Couldn't get Long Value!"
                   End If
                   
               Case Else
                   
                   MsgBox "Type Not Recognised!"
                   
           End Select
   
       Else
           MsgBox "Couldn't Query Value!"
       End If
       RegCloseKey (ll_KeyHandle)
   Else
       MsgBox "Couldn't Open Key!"
   End If
   
End Function



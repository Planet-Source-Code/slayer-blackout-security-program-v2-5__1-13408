Attribute VB_Name = "Module3"
Public Sub SaveKey(hKey As Long, strPath As String)
  Dim keyhand&
  r = RegCreateKey(hKey, strPath, keyhand&)
  r = RegCloseKey(keyhand&)
End Sub
Public Function GetString(hKey As Long, strPath As String, strValue As String)
  'EXAMPLE:
  '
  'text1.text = getstring(HKEY_CURRENT_USE
  '
  ' R, "Software\VBW\Registry", "String")
  '
  Dim keyhand As Long
  Dim datatype As Long
  Dim lResult As Long
  Dim strBuf As String
  Dim lDataBufSize As Long
  Dim intZeroPos As Integer
  r = RegOpenKey(hKey, strPath, keyhand)
  lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
  If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
      If lResult = ERROR_SUCCESS Then
        intZeroPos = InStr(strBuf, Chr$(0))
        If intZeroPos > 0 Then
          GetString = Left$(strBuf, intZeroPos - 1)
        Else
          GetString = strBuf
        End If
      End If
    End If
End Function
Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
  'EXAMPLE:
  '
  'Call savestring(HKEY_CURRENT_USER, "Sof
  '
  ' tware\VBW\Registry", "String", text1.t
  '     ex
  ' t)
  '
  Dim keyhand As Long
  Dim r As Long
  r = RegCreateKey(hKey, strPath, keyhand)
  r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
  r = RegCloseKey(keyhand)
End Sub
Function GetDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
  'EXAMPLE:
  '
  'text1.text = getdword(HKEY_CURRENT_USER
  '
  ' , "Software\VBW\Registry", "Dword")
  '
  Dim lResult As Long
  Dim lValueType As Long
  Dim lBuf As Long
  Dim lDataBufSize As Long
  Dim r As Long
  Dim keyhand As Long
  r = RegOpenKey(hKey, strPath, keyhand)
  ' Get length/data type
  lDataBufSize = 4
  lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
  If lResult = ERROR_SUCCESS Then
    If lValueType = REG_DWORD Then
      GetDWord = lBuf
    End If
    'Else
      'Call errlog("GetDWORD-" & strPath, Fals
      '
      ' e)
    'End If
  End If
     r = RegCloseKey(keyhand)
End Function
Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
  'EXAMPLE"
  '
  'Call SaveDword(HKEY_CURRENT_USER, "Soft
  '
  ' ware\VBW\Registry", "Dword", text1.tex
  '     t)
  '
  '
  Dim lResult As Long
  Dim keyhand As Long
  Dim r As Long
  r = RegCreateKey(hKey, strPath, keyhand)
  lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
  'If lResult <> error_success Then
  ' Call errlog("SetDWORD", False)
  r = RegCloseKey(keyhand)
End Function
Public Function DeleteKey(ByVal hKey As Long, ByVal strKey As String)
  'EXAMPLE:
  '
  'Call DeleteKey(HKEY_CURRENT_USER, "Soft
  '
  ' ware\VBW")
  '
  Dim r As Long
  r = RegDeleteKey(hKey, strKey)
End Function
Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
  'EXAMPLE:
  '
  'Call DeleteValue(HKEY_CURRENT_USER, "So
  '
  ' ftware\VBW\Registry", "Dword")
  '
  Dim keyhand As Long
  r = RegOpenKey(hKey, strPath, keyhand)
  r = RegDeleteValue(keyhand, strValue)
  r = RegCloseKey(keyhand)
End Function




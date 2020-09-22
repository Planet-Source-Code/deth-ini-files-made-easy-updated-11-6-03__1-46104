VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'********************************************
'    Name: INI File Functions Class Wrapper
'  Author: Lewis Miller (aka Deth)
' Purpose: Makes reading and writing to ini files a breeze
'********************************************

'API for writing to ini files
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'API for reading from ini files
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Not Used, the "Any" type cannot be passed to a vb function, so its not
'possible to wrap these calls,change them to public to use in code
Private Declare Function GetPrivateProfileStruct Lib "kernel32.dll" Alias "GetPrivateProfileStructA" (ByVal lpszSection As String, ByVal lpszKey As String, lpStruct As Any, ByVal uSizeStruct As Long, ByVal szFile As String) As Long
Private Declare Function WritePrivateProfileStruct Lib "kernel32.dll" Alias "WritePrivateProfileStructA" (ByVal lpszSection As String, ByVal lpszKey As String, lpStruct As Any, ByVal uSizeStruct As Long, ByVal szFile As String) As Long

Private mvarFilePath As String 'local copy

'filepath to ini file, this is filled in with a default path at startup
'you can change it directly or by specifying a new file path
'in a function call, any filepath used in a function call
'then becomes the default path
Public Property Let FilePath(ByVal vData As String)

    mvarFilePath = vData

End Property

Public Property Get FilePath() As String

    FilePath = mvarFilePath

End Property

'read a string from a setting in an ini file
Public Function ReadString(ByVal strSection As String, ByVal strKey As String, Optional ByVal strDefault As String, Optional ByVal strFilePath As String) As String

    On Error GoTo NoValue
    If Len(strFilePath) > 0 Then mvarFilePath = strFilePath
    ReadString = Space$(1024)
    ReadString = Left$(ReadString, GetPrivateProfileString(strSection, strKey, strDefault, ReadString, 1024&, mvarFilePath))
NoValue:

End Function

'read a number from a setting in an ini file
Public Function ReadNumber(ByVal strSection As String, ByVal strKey As String, Optional ByVal lngDefault As Long, Optional ByVal strFilePath As String) As Long

    On Error GoTo NoValue
    If Len(strFilePath) > 0 Then mvarFilePath = strFilePath
    ReadNumber = GetPrivateProfileInt(strSection, strKey, lngDefault, mvarFilePath)
NoValue:

End Function

'this is to compliment "WriteSection" and is only
'for reading a single string from an ini file
Public Function ReadSection(ByVal strSection As String, Optional ByVal strDefault As String, Optional ByVal strFilePath As String) As String

    On Error GoTo NoValue
    If Len(strFilePath) > 0 Then mvarFilePath = strFilePath
    ReadSection = Space$(1024)
    ReadSection = Left$(ReadSection, GetPrivateProfileSection(strSection, ReadSection, 1024&, mvarFilePath))
    If Len(ReadSection) = 0 Then ReadSection = strDefault
NoValue:

End Function

'this returns all the sections in an ini file,
'each name will be seperated by a Chr$(0)
Public Function ReadSectionNames(Optional ByVal strDefault As String, Optional ByVal strFilePath As String) As String

    On Error GoTo NoValue
    If Len(strFilePath) > 0 Then mvarFilePath = strFilePath
    ReadSectionNames = Space$(1024)
    ReadSectionNames = Left$(ReadSectionNames, GetPrivateProfileSectionNames(ReadSectionNames, 1024&, mvarFilePath))
    If Len(ReadSectionNames) = 0 Then ReadSectionNames = strDefault
NoValue:

End Function

'this is a handy function for reading an entire string array from disk
'that was saved with the "WriteArray" function
'each array item shouldnt be longer than 1 kb (1024) for this
'function to work, however you can modify it for longer items
'also the array to be filled must be passed byref in the arguments
Function ReadArray(ByVal strArrName As String, strArr() As String, Optional ByVal strFilePath As String) As Boolean

    On Error GoTo NoValue
  Dim lngArrCount As Long, lngCounter As Long
    If Len(strFilePath) > 0 Then mvarFilePath = strFilePath
    lngArrCount = ReadNumber(strArrName, "count")
    If lngArrCount > 0 Then
        ReDim strArr(lngArrCount - 1)
        For lngCounter = 0 To lngArrCount - 1
            strArr(lngCounter) = Space$(1024)
            strArr(lngCounter) = Left$(strArr(lngCounter), GetPrivateProfileString(strArrName, "item" & CStr(lngCounter), vbNullString, strArr(lngCounter), 1024&, mvarFilePath))
        Next lngCounter
        ReadArray = True
    End If
NoValue:

End Function

'this is same as ReadArray except for a collection
'the function will equal the new collection
Function ReadCollection(ByVal strCollName As String, Optional ByVal strFilePath As String) As Collection

  Dim lngCollCount As Long, lngCounter As Long

    If Len(strFilePath) > 0 Then mvarFilePath = strFilePath
    Set ReadCollection = New Collection
    lngCollCount = ReadNumber(strCollName, "count")
    If lngCollCount > 0 Then
        For lngCounter = 1 To lngCollCount
            ReadCollection.Add ReadString(strCollName, "item" & CStr(lngCounter))
        Next lngCounter
    End If

End Function

'write a string to an ini file
Sub WriteString(ByVal strSection As String, ByVal strKey As String, ByVal strValue As String, Optional ByVal strFilePath As String)

    If Len(strFilePath) > 0 Then mvarFilePath = strFilePath
    Call WritePrivateProfileString(strSection, strKey, strValue, mvarFilePath)

End Sub

'write a number to an ini file
Sub WriteNumber(ByVal strSection As String, ByVal strKey As String, ByVal strValue As Long, Optional ByVal strFilePath As String)

    If Len(strFilePath) > 0 Then mvarFilePath = strFilePath
    Call WritePrivateProfileString(strSection, strKey, CStr(strValue), mvarFilePath)

End Sub

'this is for writing a single string to an ini file in its own section
'example" ini.writesection "mystring",strNewString
'this can also be used to delete an entire section
Sub WriteSection(ByVal strSection As String, ByVal strValue As String, Optional ByVal strFilePath As String)

    If Len(strFilePath) > 0 Then mvarFilePath = strFilePath
    Call WritePrivateProfileSection(strSection, vbNullString, mvarFilePath)
    Call WritePrivateProfileSection(strSection, strValue, mvarFilePath)

End Sub

'this will save a string array to an ini file, which can then
'be read back into an array with "ReadArray"
Sub WriteArray(ByVal strArrName As String, Arr() As String, Optional ByVal strFilePath As String)

    On Error Resume Next
      Dim lngArrCount As Long, lngCounter As Long
        If Len(strFilePath) > 0 Then mvarFilePath = strFilePath
        
        lngArrCount = UBound(Arr)
        
        If Err.Number = 0 Then
            WriteSection strArrName, ""
            WriteNumber strArrName, "count", lngArrCount + 1
            For lngCounter = 0 To lngArrCount
                WriteString strArrName, "item" & CStr(lngCounter), Arr(lngCounter)
            Next lngCounter
          Else
            WriteNumber strArrName, "count", 0
        End If

End Sub

'this will save a collection to an ini file, collection items
'should be strings, otherwise it probably wont work very well
Sub WriteCollection(ByVal strCollName As String, colCollection As Collection, Optional ByVal strFilePath As String)

  Dim lngCollCount As Long, lngCounter As Long

    If Len(strFilePath) > 0 Then mvarFilePath = strFilePath
    If Not (colCollection Is Nothing) Then
        lngCollCount = colCollection.Count
    End If
    WriteSection strCollName, ""
    WriteNumber strCollName, "count", lngCollCount
    If lngCollCount > 0 Then
        For lngCounter = 1 To lngCollCount
            WriteString strCollName, "item" & CStr(lngCounter), colCollection(lngCounter)
        Next lngCounter
    End If

End Sub

Sub MakeIniPretty(Optional ByVal strFilePath As String)

On Error GoTo NoFile
  Dim strBuffer As String
  Dim intFileNum As Integer
  
    If Len(strFilePath) > 0 Then mvarFilePath = strFilePath
    intFileNum = FreeFile
    Open mvarFilePath For Binary As intFileNum
    strBuffer = Space$(LOF(intFileNum))
    Get #intFileNum, , strBuffer
    Close #intFileNum

    If InStr(strBuffer, vbCrLf & "[") Then
        strBuffer = Replace(strBuffer, vbCrLf & vbCrLf, vbCrLf)
        strBuffer = Replace(strBuffer, vbCrLf & "[", vbCrLf & vbCrLf & "[")
        Kill mvarFilePath
        Open mvarFilePath For Binary As #intFileNum
        Put #intFileNum, , strBuffer
        Close #intFileNum
    End If

NoFile:

End Sub

Private Sub Class_Initialize()

  'this supplys a default file path for a normal program
  'to change it, set the filepath property to whatever you wish
  'comment the line below to use this class in a dll or ocx,
  'but you must fill in the filepath property before using, to avoid errors

    mvarFilePath = App.Path & "\" & App.Title & ".ini"

End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SIJOINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' =================================================================================================================
' Class:    SIJOINI
' Author:   Steve Fitchett
' Date  :   08.16.2002
'
' Submited to SIJO Soft Corp.
' Go to http://groups.msn.com/sijosoft
' join there with your msn/hotmail passport
' through becoming a member you will abled to
' Get help and help others
' If you submit a software to SIJO Soft they will publish on net
' WITH YOUR WRITTEN PERMISSION (IF NOT PERMITTED IT CANNOT PUBLISH)
'
' A nice class wrapper around the INIFile functions
' Allows searching,deletion,modification and addition
' of Keys or Values.
'
'==================================================================================================================
'
'HOW TO USE THIS CLASS:
'
'1) TO GET VALUE :
'
'   Dim oINIFile as clsINI
'   Dim strValue as String
'
'       Set oINIFile = New clsINI
'
'       With oINIFile
'           .Path = App.Path & "\setup.ini"
'           .Section = "SectionName"
'           .Key = "KeyName"
'
'           strValue = .Value
'
'           If Not (.Success) Then
'               MsgBox "Failed to get value.", vbOkOnly + vbCritical, "Retrieving INI Value"
'           End If
'       End With
'
'2) TO WRITE VALUE :
'
'   Dim oINIFile as clsINI
'   Dim strValue as String
'
'       strValue = "TestValue"
'
'       Set oINIFile = New clsINI
'
'       With oINIFile
'           .Path = App.Path & "\setup.ini"
'           .Section = "SectionName"
'           .Key = "KeyName"
'
'           .Value = strValue
'
'           If Not (.Success) then
'               MsgBox "Failed to write value.", vbOkOnly + vbCritical, "Writing INI Value"
'           End If
'       End With
'
'3) TO GET ALL OF THE SECTIONS :
'
'   Dim oINIFile as clsINI
'   Dim strSections() as String
'   Dim lonSectionCount as Long
'   Dim lonCurrentSection as Long
'
'       Set oINIFile = New clsINI
'
'       With oINIFile
'           .Path = App.Path & "\setup.ini"
'           .EnumerateAllSections strSections(), lonSectionCount
'
'           For lonCurrentSection = 1 to lonSectionCount
'               lstINI.AddItem strSections(lonCurrentSection)
'           Next lonCurrentSection
'       End With
'
'4) TO GET ALL OF THE KEYS INSIDE OF A SECTION :
'
'   Dim oINIFile as clsINI
'   Dim strKeys() as String
'   Dim lonKeyCount as Long
'   Dim lonCurrentKey as Long
'
'       Set oINIFile = New clsINI
'
'       With oINIFile
'           .Path = App.Path & "\setup.ini"
'           .Section = "SectionName"
'           .EnumerateCurrentSection strKeys(), lonKeyCount
'
'           For lonCurrentKey = 1 To lonKeyCount
'               .Key = strKeys(lonCurrentKey)
'               listINI.AddItem .Value
'           Next loncurrentKey
'       End With
'
'5) TO DELETE A KEY :
'
'   Dim oINIFile as clsINI
'
'   Set oINIFile = New clsINI
'
'       With oINIFile
'           .Path = App.Path & "\setup.ini"
'           .Section = "SectionName"
'           .Key = "KeyName"
'           .DeleteKey
'
'           If Not (.Success) Then
'               MsgBox "Failed to delete key.", vbOkOnly + vbCritical, "Deleting INI Key"
'           End If
'       End With
'
'6) TO DELETE A SECTION :
'
'   Dim oINIFile as clsINI
'
'   Set oINIFile = new clsINI
'
'       With oINIFile
'           .Path = App.Path & "setup.ini"
'           .Section = "SectionName"
'           .DeleteSection
'
'           If Not (.Success) Then
'               MsgBox "Failed to delete section.", vbOkOnly + vbCritical, "Deleting INI Section"
'           End If
'       End With
'
'==================================================================================================================
'

Private m_strPath As String
Private m_strKey As String
Private m_strSection As String
Private m_strDefault As String
Private m_lonLastReturnCode As Long

#If Win32 Then
    
    ' Profile String functions:
    Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

#Else
    
    ' Profile String functions:
    Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

#End If

Property Get LastReturnCode() As Long
    LastReturnCode = m_lonLastReturnCode
End Property

Property Get Success() As Boolean
    Success = (m_lonLastReturnCode <> 0)
End Property

Property Let Default(strDefault As String)
    m_strDefault = strDefault
End Property
Property Get Default() As String
    Default = m_strDefault
End Property

Property Let path(strPath As String)
    m_strPath = strPath
    m_strDefault = strPath
End Property

Property Get path() As String
    path = m_strPath
End Property

Property Let Key(strKey As String)
    m_strKey = strKey
End Property

Property Get Key() As String
    Key = m_strKey
End Property

Property Let Section(strSection As String)
    m_strSection = strSection
End Property

Property Get Section() As String
    Section = m_strSection
End Property

Property Get Value() As String

Dim strBuf As String
Dim strSize As String
Dim intRetCode As Integer

    strBuf = Space$(255)
    strSize = Len(strBuf)
    intRetCode = GetPrivateProfileString(Section, Key, Default, strBuf, strSize, path)
    
    If (strSize > 0) Then
        
        Value = Left$(strBuf, intRetCode)
    
    Else
        
        Value = ""
    
    End If

End Property

Property Let Value(strValue As String)

Dim intPos As Integer
    
    ' Strip chr$(0):
    intPos = InStr(strValue, Chr$(0))
    
    Do While intPos <> 0
        
        strValue = Left$(strValue, (intPos - 1)) & Mid$(strValue, (intPos + 1))
        intPos = InStr(strValue, Chr$(0))
    
    Loop
    
    m_lonLastReturnCode = WritePrivateProfileString(Section, Key, strValue, path)

End Property

Public Sub DeleteKey()
    
    m_lonLastReturnCode = WritePrivateProfileString(Section, Key, 0&, path)

End Sub

Public Sub DeleteSection()
    
    m_lonLastReturnCode = WritePrivateProfileString(Section, 0&, 0&, path)

End Sub

Property Get INISection() As String

Dim strBuf As String
Dim strSize As String
Dim intRetCode As Integer

    strBuf = Space$(8192)
    strSize = Len(strBuf)
    intRetCode = GetPrivateProfileString(Section, 0&, Default, strBuf, strSize, path)
    
    If (strSize > 0) Then
        
        INISection = Left$(strBuf, intRetCode)
    
    Else
        
        INISection = ""
    
    End If

End Property

Property Let INISection(strSection As String)
    
    m_lonLastReturnCode = WritePrivateProfileString(Section, 0&, strSection, path)

End Property

Property Get Sections() As String

Dim strBuf As String
Dim strSize As String
Dim intRetCode As Integer

    strBuf = Space$(8192)
    strSize = Len(strBuf)
    intRetCode = GetPrivateProfileString(0&, 0&, Default, strBuf, strSize, path)
    
    If (strSize > 0) Then
        
        Sections = Left$(strBuf, intRetCode)
    
    Else
        
        Sections = ""
    
    End If

End Property

Public Sub EnumerateCurrentSection(ByRef sKey() As String, ByRef intCount As Long)

Dim strSection As String
Dim intPos As Long
Dim intNextPos As Long
Dim strCur As String
    
    intCount = 0
    Erase sKey
    strSection = INISection
    
    If (Len(strSection) > 0) Then
        
        intPos = 1
        intNextPos = InStr(intPos, strSection, Chr$(0))
        
        Do While intNextPos <> 0
            
            strCur = Mid$(strSection, intPos, (intNextPos - intPos))
            
            If (strCur <> Chr$(0)) Then
                
                intCount = intCount + 1
                
                ReDim Preserve sKey(1 To intCount) As String
                
                sKey(intCount) = Mid$(strSection, intPos, (intNextPos - intPos))
                intPos = intNextPos + 1
                intNextPos = InStr(intPos, strSection, Chr$(0))
            
            End If
        
        Loop
    
    End If

End Sub

Public Sub EnumerateAllSections(ByRef strSections() As String, ByRef intCount As Long)

Dim strIniFile As String
Dim intPos As Long
Dim intNextPos As Long
Dim strCur As String
    
    intCount = 0
    Erase strSections
    strIniFile = Sections
    
    If (Len(strIniFile) > 0) Then
        
        intPos = 1
        intNextPos = InStr(intPos, strIniFile, Chr$(0))
        
        Do While intNextPos <> 0
            
            If (intNextPos <> intPos) Then
                
                strCur = Mid$(strIniFile, intPos, (intNextPos - intPos))
                intCount = intCount + 1
                
                ReDim Preserve strSections(1 To intCount) As String
                
                strSections(intCount) = strCur
            
            End If
            
            intPos = intNextPos + 1
            intNextPos = InStr(intPos, strIniFile, Chr$(0))
        
        Loop
    
    End If

End Sub


''Option Explicit
''' =========================================================
''' Class:    clsINI
''' Author:   Steve Fitchett
''' Date  :   08.16.2002
'''
''' A nice class wrapper around the INIFile functions
''' Allows searching,deletion,modification and addition
''' of Keys or Values.
''
''Private m_strPath As String
'''Private m_strKey As String
'''Private m_strSection As String
''Private m_strDefault As String
''Private m_lonLastReturnCode As Long
''
''#If Win32 Then
''    ' Profile String functions:
''    Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
''    Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
''#Else
''    ' Profile String functions:
''    Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
''    Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
''#End If
''
''Property Get LastReturnCode() As Long
''    LastReturnCode = m_lonLastReturnCode
''End Property
''
''Property Get Success() As Boolean
''    Success = (m_lonLastReturnCode <> 0)
''End Property
''
''Property Let Default(strDefault As String)
''    m_strDefault = strDefault
''End Property
''Property Get Default() As String
''    Default = m_strDefault
''End Property
''
''Property Let Path(strPath As String)
''    m_strPath = strPath
''End Property
''
''Property Get Path() As String
''    Path = m_strPath
''End Property
''
'''Property Let Key(strKey As String)
'''    m_strKey = strKey
'''End Property
'''
'''Property Get Key() As String
'''    Key = m_strKey
'''End Property
'''
'''Property Let Section(strSection As String)
'''    m_strSection = strSection
'''End Property
'''
'''Property Get Section() As String
'''    Section = m_strSection
'''End Property
''
''Property Get Value(strSection As String, strKey As String) As String
''
''Dim strBuf As String
''Dim strSize As String
''Dim intRetCode As Integer
''
''    strBuf = Space$(255)
''    strSize = Len(strBuf)
''    intRetCode = GetPrivateProfileString(strSection, strKey, Default, strBuf, strSize, Path)
''
''    If (strSize > 0) Then
''
''        Value = Left$(strBuf, intRetCode)
''
''    Else
''
''        Value = ""
''
''    End If
''
''End Property
''
''Property Let Value(strSection As String, strKey As String, strValue As String)
''
''Dim intPos As Integer
''
''    ' Strip chr$(0):
''    intPos = InStr(strValue, Chr$(0))
''
''    Do While intPos <> 0
''
''        strValue = Left$(strValue, (intPos - 1)) & Mid$(strValue, (intPos + 1))
''        intPos = InStr(strValue, Chr$(0))
''
''    Loop
''
''    m_lonLastReturnCode = WritePrivateProfileString(strSection, strKey, strValue, Path)
''
''End Property
''
''Public Sub DeleteKey(strSection As String, strKey As String)
''
''    m_lonLastReturnCode = WritePrivateProfileString(strSection, strKey, 0&, Path)
''
''End Sub
''
''Public Sub DeleteSection(strSection As String)
''
''    m_lonLastReturnCode = WritePrivateProfileString(strSection, 0&, 0&, Path)
''
''End Sub
''
''Property Get INISection(strSection As String) As String
''
''Dim strBuf As String
''Dim strSize As String
''Dim intRetCode As Integer
''
''    strBuf = Space$(8192)
''    strSize = Len(strBuf)
''    intRetCode = GetPrivateProfileString(strSection, 0&, Default, strBuf, strSize, Path)
''
''    If (strSize > 0) Then
''
''        INISection = Left$(strBuf, intRetCode)
''
''    Else
''
''        INISection = ""
''
''    End If
''
''End Property
'''
'''Property Let INISection(strSection As String)
'''
'''    m_lonLastReturnCode = WritePrivateProfileString(strSection, 0&, strSection, Path)
'''
'''End Property
''
''Property Get Sections() As String
''
''Dim strBuf As String
''Dim strSize As String
''Dim intRetCode As Integer
''
''    strBuf = Space$(8192)
''    strSize = Len(strBuf)
''    intRetCode = GetPrivateProfileString(0&, 0&, Default, strBuf, strSize, Path)
''
''    If (strSize > 0) Then
''
''        Sections = Left$(strBuf, intRetCode)
''
''    Else
''
''        Sections = ""
''
''    End If
''
''End Property
''
'''Public Sub EnumerateCurrentSection(ByRef sKey() As String, ByRef intCount As Long)
'''
'''Dim strSection As String
'''Dim intPos As Long
'''Dim intNextPos As Long
'''Dim strCur As String
'''
'''    intCount = 0
'''    Erase sKey
''''    strSection = Section
'''
'''    If (Len(strSection) > 0) Then
'''
'''        intPos = 1
'''        intNextPos = InStr(intPos, strSection, Chr$(0))
'''
'''        Do While intNextPos <> 0
'''
'''            strCur = Mid$(strSection, intPos, (intNextPos - intPos))
'''
'''            If (strCur <> Chr$(0)) Then
'''
'''                intCount = intCount + 1
'''
'''                ReDim Preserve sKey(1 To intCount) As String
'''
'''                sKey(intCount) = Mid$(strSection, intPos, (intNextPos - intPos))
'''                intPos = intNextPos + 1
'''                intNextPos = InStr(intPos, strSection, Chr$(0))
'''
'''            End If
'''
'''        Loop
'''
'''    End If
'''
'''End Sub
'''
'''Public Sub EnumerateAllSections(ByRef strSections() As String, ByRef intCount As Long)
'''
'''Dim strIniFile As String
'''Dim intPos As Long
'''Dim intNextPos As Long
'''Dim strCur As String
'''
'''    intCount = 0
'''    Erase strSections
'''    strIniFile = Sections
'''
'''    If (Len(strIniFile) > 0) Then
'''
'''        intPos = 1
'''        intNextPos = InStr(intPos, strIniFile, Chr$(0))
'''
'''        Do While intNextPos <> 0
'''
'''            If (intNextPos <> intPos) Then
'''
'''                strCur = Mid$(strIniFile, intPos, (intNextPos - intPos))
'''                intCount = intCount + 1
'''
'''                ReDim Preserve strSections(1 To intCount) As String
'''
'''                strSections(intCount) = strCur
'''
'''            End If
'''
'''            intPos = intNextPos + 1
'''            intNextPos = InStr(intPos, strIniFile, Chr$(0))
'''
'''        Loop
'''
'''    End If
'''
'''End Sub
''
''
''Private Sub Class_Initialize()
''
''    Path = App.Path & "\setup.ini"
''
''End Sub


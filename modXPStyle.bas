Attribute VB_Name = "XP_Style"
'------------------------------------------------------------
' Module    : modXPStyle.bas
' Purpose   : This routine will automatically write a small XML file
'             (manifest) in your program folder. Which carries some
'             info about your program and let Windows XP draw your
'             application. It won't be affected by the exe name
' Author    : Voodoo Attack!!
'             E-Mail: voodooattack@hotmail.com
' Source    : http://www.thescarms.com/vbasic/XPStyle.asp
' Updated ..: 09/28/2005 - Robert Edwards
'                          Added InIDE routine
'------------------------------------------------------------
'
Option Explicit

Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

'------------------------------------------------------------
' Purpose   : Creates the manifest file
' Arguments : AutoRestart - True will restart the exe, false will let it run as is
'             CreateNew   - True will re-create the file even if one already exists
' Returns   : True = the manifest file does NOT exist
'             False = the manifest file does exist
'------------------------------------------------------------
'
Public Function XPStyle(Optional AutoRestart As Boolean = True, _
                        Optional CreateNew As Boolean) As Boolean

    InitCommonControls
    
    On Error Resume Next

    Dim XML             As String
    Dim ManifestCheck   As String
    Dim strManifest     As String
    Dim FreeFileNo      As Integer

    If InIDE() Then
        XPStyle = False
        Exit Function
    End If
    
    If AutoRestart = True Then CreateNew = False

    '(put the XML in a string)
    XML = ("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?> " & vbCrLf & "<assembly " & vbCrLf & _
           "   xmlns=""urn:schemas-microsoft-com:asm.v1"" " & vbCrLf & _
           "   manifestVersion=""1.0"">" & vbCrLf & _
           "<assemblyIdentity " & vbCrLf & _
           "    processorArchitecture=""x86"" " & vbCrLf & _
           "    version=""EXEVERSION""" & vbCrLf & _
           "    type=""win32""" & vbCrLf & _
           "    name=""EXENAME""/>" & vbCrLf & _
           "    <description>EXEDESCRIBTION</description>" & vbCrLf & _
           "    <dependency>" & vbCrLf & _
           "    <dependentAssembly>" & vbCrLf & _
           "    <assemblyIdentity" & vbCrLf & _
           "         type=""win32""" & vbCrLf & _
           "         name=""Microsoft.Windows.Common-Controls""" & vbCrLf & _
           "         version=""6.0.0.0""" & vbCrLf & _
           "         publicKeyToken=""6595b64144ccf1df""" & vbCrLf & _
           "         language=""*""" & vbCrLf & _
           "         processorArchitecture=""x86""/>" & vbCrLf & _
           "    </dependentAssembly>" & vbCrLf & _
           "    </dependency>" & vbCrLf & _
           "</assembly>" & vbCrLf & "")

    '--- Set the name of the manifest
    strManifest = App.Path & "\" & App.EXEName & ".exe.manifest"
    
    '--- Check the app manifest file
    ManifestCheck = Dir(strManifest, vbNormal + vbSystem + vbHidden + vbReadOnly + vbArchive)
    
    '--- If not found.. make a new one
    If ManifestCheck = "" Or CreateNew = True Then
        
        '--- Replaces the string "EXENAME" with the program's exe file name
        XML = Replace(XML, "EXENAME", App.EXEName & ".exe")
        
        '--- Replaces the "EXEVERSION" string
        XML = Replace(XML, "EXEVERSION", App.Major & "." & App.Minor & "." & App.Revision & ".0")
        
        '--- Replaces the app describtion
        XML = Replace(XML, "EXEDESCRIBTION", App.FileDescription)
        
        FreeFileNo = FreeFile  'get the next avilabel file
        
        If ManifestCheck <> "" Then
            SetAttr strManifest, vbNormal
            Kill strManifest
        End If
    
        Open strManifest For Binary As #(FreeFileNo) 'open the file
        Put #(FreeFileNo), , XML        'uses 'put' to set the file content.. note that 'put' (binary mode) is much faster than 'print'(output mode)
        Close #(FreeFileNo)             'close the file
    
        SetAttr strManifest, vbHidden + vbSystem
        
        If ManifestCheck = "" Then
            XPStyle = False             'return false.. this means that the manifest file does not exist
        Else
            XPStyle = True
        End If
    
        '--- If in automode (default), the program will restart
        If AutoRestart = True Then
            Shell App.Path & "\" & App.EXEName & ".exe " & Command$, vbNormalFocus  'restart the program and include command line parameters (if any)
            End                         'end the session
        End If
    Else                                'the manifest file exists
        XPStyle = True                  'return true
    End If

End Function

'------------------------------------------------------------
' Purpose   : Checks to see if we are inside the IDE
' Arguments : none
' Returns   : True if inside vb's ide
'             False if not inside vb's ide
'------------------------------------------------------------
'
Private Function InIDE() As Boolean
    
    On Error GoTo Err_Handler
    
    InIDE = False
    Debug.Print 1 / 0   'will generate an error inside the IDE
    Exit Function

'==========
Err_Handler:
'==========
    InIDE = True
    Exit Function

End Function

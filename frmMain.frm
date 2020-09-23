VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Software Scanner"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Exclude"
      Height          =   972
      Left            =   3960
      TabIndex        =   7
      Top             =   0
      Width           =   1572
      Begin VB.CheckBox ckHotfix 
         Caption         =   "Hotfix(s)"
         Height          =   252
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1212
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Include"
      Height          =   972
      Left            =   2280
      TabIndex        =   6
      Top             =   0
      Width           =   1572
      Begin VB.CheckBox ckPrinters 
         Caption         =   "Printer(s)"
         Height          =   252
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1212
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   372
      Left            =   5760
      TabIndex        =   4
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Default         =   -1  'True
      Height          =   372
      Left            =   5760
      TabIndex        =   5
      Top             =   600
      Width           =   852
   End
   Begin VB.ListBox lstApps 
      Height          =   3570
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   6495
   End
   Begin VB.TextBox txtHost 
      Height          =   288
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblApps 
      Caption         =   "Installed Application(s)"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1932
   End
   Begin VB.Label lblHost 
      Caption         =   "Host Name"
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1932
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------
' Module    : frmMain.frm
' Purpose   : Main module
' Author    : R. Edwards
'------------------------------------------------------------
'
Option Explicit

Dim AppList() As String     'array holding app list
Dim Auto      As Boolean    'auto mode
Dim HostName  As String     'host name

'------------------------------------------------------------
' Purpose   : Exclude 'Hotfixs' from list
' Arguments : none
'------------------------------------------------------------
'
Private Sub ckHotfix_Click()
    
    '--- Retreive installed apps
    If GetApps Then
        DisplayPrograms  'update displayed list of apps
    End If

End Sub

'------------------------------------------------------------
' Purpose   : When the user clicks the 'Exit' button, the form unloads
' Arguments : none
'------------------------------------------------------------
'
Private Sub cmdExit_Click()
    
    Unload Me

End Sub

'------------------------------------------------------------
' Purpose   : When the user clicks the 'Save' button, the
'             list of installed applications is saved to a
'             text file using the HostName for the filename
' Arguments : none
'------------------------------------------------------------
'
Private Sub cmdSave_Click()

    Dim iCount   As Integer
    Dim FileName As String
    Dim fn       As Long
    Dim p        As Printer
    
    On Error GoTo ErrRoutine

    fn = FreeFile()
    
    Open App.Path & "\" & HostName & ".txt" For Output As #fn

    If Auto Then
        Print #fn, App.Title & " - (Automode)"
    Else
        Print #fn, App.Title
    End If
    Print #fn, "==================================================="
    Print #fn, "HostName: " & HostName
    Print #fn, "---------------------------------------------------"
    Print #fn, "Installed Application(s) as of: " & Now()
    Print #fn, "==================================================="
    
    For iCount = 1 To UBound(AppList)
        Print #fn, Format(iCount, "000") & " - " & AppList(iCount) 'add entry to printed list
    Next iCount
    
    '--- Include a list of installed printer if selected
    If Me.ckPrinters Or Auto Then
        Print #fn, "==================================================="
        Print #fn, "Installed Printer(s) as of: " & Now()
        Print #fn, "---------------------------------------------------"
        For Each p In Printers
            Print #fn, p.DeviceName
        Next p
    End If
    
    Print #fn, "==================================================="
    Close #fn
    
'==========
ExitRoutine:
'==========
    Exit Sub

'==========
ErrRoutine:
'==========
    MsgBox "Unable to Save Report on " & App.Path, vbExclamation, App.Title

End Sub

'------------------------------------------------------------
' Purpose   : Event run when form is initialized
' Arguments : none
'------------------------------------------------------------
'
Private Sub Form_Initialize()
    
    '--- Create the manifest file w/ auto restart
    XPStyle True, False

End Sub

'------------------------------------------------------------
' Purpose   : Events runs when form is loaded
' Arguments : none
'------------------------------------------------------------
'
Private Sub Form_Load()
    
    '--- Initilization
    App.Title = "Software Scanner v" & App.Major & "." & Format(App.Minor, "00")
    Me.Caption = App.Title
    Me.cmdSave.Enabled = False
    
    '--- Center form on scrren
    SetFormPosition Me, (JMScreenHeight - Me.Height) / 2, (JMScreenWidth - Me.Width) / 2
    
    '--- Are we in 'Auto' mode
    If LCase(Command()) = "/auto" Then
        Auto = True
    End If
    
    '--- Return computer name
    HostName = ComputerName
    Me.txtHost.Text = HostName

    '--- Retreive installed apps
    If GetApps Then
        DisplayPrograms             'display list of apps
    End If

    '--- If we're in 'Auto' mode, save data and exit
    If Auto Then
        ckHotfix.Value = vbUnchecked
        cmdSave_Click
        cmdExit_Click
    End If

End Sub

'------------------------------------------------------------
' Purpose   : Events runs when form is unloaded
' Arguments : none
'------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)
    
    End

End Sub

'------------------------------------------------------------
' Purpose   : Retrives installed applications from the registry
'             and loads the names into the global array AppList
' Arguments : none
'------------------------------------------------------------
'
Private Function GetApps() As Boolean
    
    Dim iCount As Integer
    Dim returnName As Collection
    Dim returnSubs As Collection
    Dim DisplayName As String
    Dim InstallDate As String
    Dim UninstallString As String
    Dim Version As String
    Dim FinalCount As Integer
    
    On Error GoTo ErrRoutine
    
    Call EnumRegKeys(returnName, returnSubs, "HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")

    If returnName.Count > 0 Then
        
        ReDim Preserve AppList(returnName.Count) As String
        
        For iCount = 1 To returnName.Count
            DisplayName = GetSetting("", returnName(iCount), "DisplayName", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            InstallDate = GetSetting("", returnName(iCount), "Installdate", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            UninstallString = GetSetting("", returnName(iCount), "UninstallString", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            Version = GetSetting("", returnName(iCount), "Version", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            If DisplayName <> "" And (UninstallString <> "" Or Version <> "") Then
                
                '--- If exclude hotfixs checked
                If ckHotfix.Value = vbChecked Then
                    
                    If (InStr(1, LCase(DisplayName), "hotfix") = 0) And (InStr(1, LCase(DisplayName), "kb") = 0) Then
                        FinalCount = FinalCount + 1
                        If InstallDate <> "" Then
                            If Len(InstallDate) = 8 Then
                                DisplayName = DisplayName & " - " & Mid(InstallDate, 3, 2) & "/" & Mid(InstallDate, 7, 2) & "/" & Mid(InstallDate, 1, 4)
                            End If
                        End If
                        AppList(FinalCount) = DisplayName 'add to array
                    End If
                
                '--- Else, include hotfixs
                Else
                    FinalCount = FinalCount + 1
                    If InstallDate <> "" Then
                        If Len(InstallDate) = 8 Then
                            DisplayName = DisplayName & " - " & Mid(InstallDate, 3, 2) & "/" & Mid(InstallDate, 7, 2) & "/" & Mid(InstallDate, 1, 4)
                        End If
                    End If
                    AppList(FinalCount) = DisplayName 'add to array
                End If
            End If
        Next iCount
    End If
    
    ReDim Preserve AppList(FinalCount) As String
    
    '--- Sort list
    QuickSortMe AppList()
    
    If FinalCount > 0 Then
        Me.cmdSave.Enabled = True
        GetApps = True
    End If

'==========
ExitRoutine:
'==========
    Exit Function
    
'==========
ErrRoutine:
'==========
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetPrograms of Form frmMain"

End Function

'------------------------------------------------------------
' Purpose   : Displays the list of installed applications
' Arguments : none
'------------------------------------------------------------
'
Private Sub DisplayPrograms()

    Dim iCount As Integer
    
    On Error GoTo ErrRoutine

    lstApps.Clear
    
    For iCount = 1 To UBound(AppList)
        Me.lstApps.AddItem Format(iCount, "000") & " - " & (AppList(iCount)) 'add entry to display list
    Next iCount
        
'==========
ExitRoutine:
'==========
    Exit Sub

'==========
ErrRoutine:
'==========
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PrintPrograms of Form frmMain"

End Sub

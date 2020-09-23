VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00905747&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Self Register"
   ClientHeight    =   5430
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7395
   FillColor       =   &H00000001&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUnRegAll 
      BackColor       =   &H00905747&
      Caption         =   "U&nregister All"
      Height          =   375
      Left            =   5610
      TabIndex        =   10
      Top             =   3060
      Width           =   1665
   End
   Begin MSComctlLib.ProgressBar pbBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   13
      Top             =   5085
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   609
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdUnRegSelection 
      Appearance      =   0  'Flat
      BackColor       =   &H00905747&
      Caption         =   "Unr&egister Selection"
      Height          =   375
      Left            =   5610
      TabIndex        =   8
      Top             =   2145
      Width           =   1665
   End
   Begin VB.CommandButton cmdRegSelection 
      Appearance      =   0  'Flat
      BackColor       =   &H00905747&
      Caption         =   "&Register Selection"
      Height          =   375
      Left            =   5610
      TabIndex        =   7
      Top             =   1785
      Width           =   1665
   End
   Begin VB.CommandButton cmdUnselectAll 
      Appearance      =   0  'Flat
      BackColor       =   &H00905747&
      Caption         =   "&Unselect All"
      Height          =   375
      Left            =   5610
      TabIndex        =   6
      Top             =   1230
      Width           =   1665
   End
   Begin VB.CommandButton cmdSelectAll 
      Appearance      =   0  'Flat
      BackColor       =   &H00905747&
      Caption         =   "&Select All"
      Height          =   375
      Left            =   5610
      TabIndex        =   5
      Top             =   870
      Width           =   1665
   End
   Begin VB.CommandButton cmdListFiles 
      Caption         =   "List Files"
      Height          =   375
      Left            =   5610
      TabIndex        =   2
      Top             =   300
      Width           =   1665
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F76
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5145
      TabIndex        =   1
      ToolTipText     =   "Browse folders"
      Top             =   300
      Width           =   345
   End
   Begin VB.CommandButton cmdRegAll 
      Appearance      =   0  'Flat
      BackColor       =   &H00905747&
      Caption         =   "Register &All"
      Height          =   375
      Left            =   5610
      TabIndex        =   9
      Top             =   2700
      Width           =   1665
   End
   Begin VB.TextBox txtComServerDir 
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   300
      Width           =   4965
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4095
      Left            =   90
      TabIndex        =   4
      Top             =   870
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Status"
         Object.Width           =   8414
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   4095
      Left            =   90
      TabIndex        =   3
      Top             =   870
      Visible         =   0   'False
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Status"
         Object.Width           =   8431
      EndProperty
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6420
      TabIndex        =   16
      Top             =   4260
      Width           =   870
   End
   Begin VB.Label lblBy 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Johan van Rensburg"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5310
      TabIndex        =   15
      Top             =   4500
      Width           =   1980
   End
   Begin VB.Label lblEmail 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "email : johanjvr@y.co.za"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5145
      TabIndex        =   14
      Top             =   4710
      Width           =   2145
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   630
      Width           =   270
   End
   Begin VB.Label lblReg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Path"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   90
      Width           =   330
   End
   Begin VB.Menu mExit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu mOptions 
      Caption         =   "&Options"
      Begin VB.Menu mTypes 
         Caption         =   "Register Type's"
         Begin VB.Menu mOcx 
            Caption         =   "&Ocx's"
            Checked         =   -1  'True
         End
         Begin VB.Menu mDll 
            Caption         =   "&Dll's"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mLog 
         Caption         =   "&Application Log"
      End
   End
   Begin VB.Menu mAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oDll As New clsSelfReg
Private oFolder As Shell32.Folder
Private strComServers As String
Private bRegister As Boolean, bListWasSelected As Boolean

Private Sub cmdListFiles_Click()
    
    bListWasSelected = True
    
    'Get a list of all the folders in the that was selected
    AddFilesToListView
    
    ListView2.Visible = True
    ListView1.Visible = False

End Sub

Private Sub cmdRegAll_Click()

    bRegister = True
    ListView1.ListItems.Clear
    Screen.MousePointer = vbHourglass
    UpdateList
    Screen.MousePointer = Default
    ListView2.Visible = False
    ListView1.Visible = True

End Sub

Private Sub cmdBrowse_Click()
    
    'Call the BrowseFolder Method
    txtComServerDir = BrowseFolder

End Sub

Private Sub cmdRegSelection_Click()
        
    If ListView2.ListItems.Count = 0 Then Exit Sub
        
    bRegister = True
    ListView1.ListItems.Clear
    Screen.MousePointer = vbHourglass
    UpdateList True
    Screen.MousePointer = Default
    ListView2.Visible = False
    ListView1.Visible = True

End Sub

Private Sub cmdSelectAll_Click()
    Dim i As Integer
    
    'Check all item in listview
    For i = 1 To ListView2.ListItems.Count
        ListView2.ListItems.Item(i).Checked = True 'IIf(ListView1.ListItems.Item(I).Checked, True, False)
    Next i
    
End Sub

Private Sub cmdUnRegAll_Click()

    'Update Internal var
    bRegister = False
    
    'clear the listview
    ListView1.ListItems.Clear
    
    'update the Mouse pointer
    Screen.MousePointer = vbHourglass
    
    'call the UpdateList method
    UpdateList
    
    'restore the Mouse Pointer
    Screen.MousePointer = Default
    
    'update visibility
    ListView2.Visible = False
    ListView1.Visible = True

End Sub

Private Sub cmdUnRegSelection_Click()
        
    'If there is not item then exit
    If ListView2.ListItems.Count = 0 Then Exit Sub
    
    'Update Internal var
    bRegister = False
    
    'Clear the ListView to add to
    ListView1.ListItems.Clear
    
    'Update the Mouse Pointer to an Hourglass
    Screen.MousePointer = vbHourglass
    
    'Call the UpdateList Method
    UpdateList True
    
    'Restore the Mouse Pointer to default arrow
    Screen.MousePointer = Default
    
    'Change the Visibility
    ListView2.Visible = False
    ListView1.Visible = True

End Sub

Private Sub cmdUnselectAll_Click()

    Dim i As Integer
    
    'Uncheck all the files
    For i = 1 To ListView2.ListItems.Count
        ListView2.ListItems.Item(i).Checked = False 'IIf(ListView1.ListItems.Item(I).Checked, True, False)
    Next i

End Sub

Private Sub Form_Load()
    
    'When the Form load load settings
    DoEvents
    
    Set oDll = New clsSelfReg
    
    'Get the Path from the Registry
    txtComServerDir = GetSetting("SelfReg", "Config", "LastPath", "")
    txtComServerDir.SelStart = Len(txtComServerDir)
    
    'Change visibilitie
    ListView1.Visible = False
    ListView2.Visible = True
    
    'Update the Version Lable
    lblVersion = "Version : " & App.Major & "." & App.Revision

End Sub

Private Function BrowseFolder() As String

On Error GoTo ErrHandler
    Dim strRootPath As String
    Dim oShell As New Shell32.Shell
    Dim strPath As String
    Dim i As Integer
    
    'Default Starting Path
    strRootPath = "My Computer" 'Trim(txtComServerDir)
    
    'Call the BrowseForFolder Method
    Set oFolder = oShell.BrowseForFolder(Me.hWnd, _
        "Directory to Use", _
        0, strRootPath)
    
    'Conditions is not meat exit
    If oFolder Is Nothing Then Exit Function
    If oFolder.Items Is Nothing Then Exit Function
    If oFolder.Items.Count = 0 Then Exit Function
    
    'Compile the Return Path
    strPath = oFolder.Items.Item(0).Path
    strPath = Left(strPath, InStrRev(strPath, "\"))
    
    If Left(strPath, 2) = "::" Then
        strRootPath = oShell.NameSpace(oFolder)
    Else
        strRootPath = strPath
    End If
    
    'Return the Path
    BrowseFolder = Left(strRootPath, Len(strRootPath) - 1)
    
    Exit Function
ErrHandler:
    Stop
    MsgBox Err.Description
    
End Function

Public Sub UpdateList(Optional bSelectionOnly As Boolean)

On Error GoTo ErrHandler
    Dim tCount As Long
    Dim oShell As New Shell32.Shell
    Dim strPath As String
    Dim i As Integer
    Dim bOk As Boolean
    
    'Log Actions
    frmLog.txtLog = frmLog.txtLog + "Started : " & CStr(Now) + vbCrLf
    
    'Exit if Path is empty
    If Len(Trim(txtComServerDir)) = 0 Then Exit Sub
    
    'Get list of file in folder
    Set oFolder = oShell.NameSpace(txtComServerDir + "\")
    
    'Update the Progressbar
    pbBar.Max = oFolder.Items.Count
    pbBar.Value = 0
    
    For i = 0 To oFolder.Items.Count - 1
        Dim F As Shell32.ShellFolderItem
        Set F = oFolder.Items.Item(i)
        
        'Update the Progress Bar
        If pbBar.Value <= pbBar.Max Then pbBar.Value = pbBar.Value + 1
        
        'if selection only then was clicked then check if current item is in the list and was selected
        If bSelectionOnly Then
            Dim oItem As ListItem
            Set oItem = ListView2.FindItem(F.Name)
            If Not oItem Is Nothing Then If oItem.Checked Then bOk = True
        Else
            bOk = True
        End If
        
        'Register the Current DLL, OCX
        If (UCase(Right(F.Name, 3)) = UCase("ocx") And mOcx.Checked And bOk) Or _
           (UCase(Right(F.Name, 3)) = UCase("dll") And mDll.Checked And bOk) Then
            
            'Log Actions
            frmLog.txtLog = frmLog.txtLog + F.Path + vbCrLf
            
            'Do the Registration Now
            oDll.RegisterServer Me.hWnd, F.Path, bRegister
        End If
        
        bOk = False
    Next i
    
    'Log Actions
    frmLog.txtLog = frmLog.txtLog + "Ended : " & CStr(Now) + vbCrLf
    
    Exit Sub
ErrHandler:
    'Stop
    MsgBox Err.Description
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  'Unload the Log Form from Memory
  Unload frmLog
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Save the Setting to the Registry
    SaveSetting "SelfReg", "Config", "LastPath", txtComServerDir

End Sub

Private Sub mAbout_Click()
    
    'Load the About Form
    Load frmAbout
    
    'Display the About UI
    frmAbout.Show vbModal

End Sub

Private Sub mDll_Click()
    
    'Update the Status
    If mDll.Checked Then
        mDll.Checked = False
    Else
        mDll.Checked = True
    End If

End Sub

Private Sub mExit_Click()
    
    'Save the Setting and Exit Application
    SaveSetting "SelfReg", "Config", "LastPath", txtComServerDir
    
    'End Application
    End

End Sub

Private Sub mLog_Click()
    
    'Load the Log UI into memory
    Load frmLog
    
    'Display the Log UI
    frmLog.Show vbModal

End Sub

Private Sub mOcx_Click()
    
    'Update the Status
    If mOcx.Checked Then
        mOcx.Checked = False
    Else
        mOcx.Checked = True
    End If

End Sub

Private Sub AddFilesToListView()
    Dim oShell As New Shell32.Shell
    Dim i As Integer
    
    'If the Path is empty don't continue
    If Len(Trim(txtComServerDir)) = 0 Then Exit Sub
    
    'Clear the ListView
    ListView2.ListItems.Clear
    
    'Get a Folder Object of all the Files
    Set oFolder = oShell.NameSpace(txtComServerDir + "\")
    
    'Update the progress bar
    pbBar.Max = oFolder.Items.Count
    pbBar.Value = 0
    
    'Loop through the oFolder Object
    For i = 0 To oFolder.Items.Count - 1
        Dim F As Shell32.ShellFolderItem
        Set F = oFolder.Items.Item(i)
        
        'Update the Progress Bar
        If pbBar.Value <= pbBar.Max Then pbBar.Value = pbBar.Value + 1
        
        'Add Files to the ListView
        If (UCase(Right(F.Name, 3)) = UCase("ocx") And mOcx.Checked) Or _
           (UCase(Right(F.Name, 3)) = UCase("dll") And mDll.Checked) Then
            
            ListView2.ListItems.Add , , F.Name, , 3
        End If
        
    'Move to next File in oFolder
    Next i
    
    'Exit before the Error Handler
    Exit Sub
ErrHandler:
    'Display Error Message
    MsgBox Err.Description
End Sub


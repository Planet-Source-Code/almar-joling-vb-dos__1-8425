Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Private Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleOutput As Long, dwMode As Long) As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&
Private Const FOREGROUND_RED = &H4
Private Const FOREGROUND_GREEN = &H2
Private Const FOREGROUND_BLUE = &H1
Private Const FOREGROUND_INTENSITY = &H8
Private Const BACKGROUND_RED = &H40
Private Const BACKGROUND_GREEN = &H20
Private Const BACKGROUND_BLUE = &H10
Private Const BACKGROUND_INTENSITY = &H80
Private Const ENABLE_LINE_INPUT = &H2
Private Const ENABLE_ECHO_INPUT = &H4
Private Const ENABLE_MOUSE_INPUT = &H10
Private Const ENABLE_PROCESSED_INPUT = &H1
Private Const ENABLE_WINDOW_INPUT = &H8
Private Const ENABLE_PROCESSED_OUTPUT = &H1
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2
Private hConsoleIn As Long
Private hConsoleOut As Long
Private hConsoleErr As Long
Private Declare Function SetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Private Declare Function SetConsoleCursorPosition Lib "kernel32" (ByVal hConsoleOutput As Long, dwCursorPosition As COORD) As Long
Private Type CONSOLE_CURSOR_INFO
        dwSize As Long
        bVisible As Long
End Type
Private Type COORD
        X As Integer
        y As Integer
End Type
Private Declare Function GetConsoleCP Lib "kernel32" () As Long
Private Const Spacing = 40

Private Function CGet() As String
Dim sUserInput As String * 256
Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), vbNull, vbNull)
Form1.Label1.Caption = Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)
CGet = Form1.Label1.Caption 'Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)
End Function
Private Sub CPrint(Cout As String)
WriteConsole hConsoleOut, Cout, Len(Cout), vbNull, vbNull
End Sub
Sub Main()
Dim UserIn As String
Dim Receive As String
AllocConsole
SetConsoleTitle "VB Dos"
hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)
SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED Or FOREGROUND_GREEN Or FOREGROUND_BLUE Or FOREGROUND_INTENSITY Or BACKGROUND_GREEN
CPrint "VB Dos 1.0" & vbCrLf
SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED Or FOREGROUND_GREEN Or FOREGROUND_BLUE Or FOREGROUND_INTENSITY Or BACKGROUND_RED
CPrint "Loaded." & vbCrLf
SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED Or FOREGROUND_GREEN Or FOREGROUND_BLUE
'CPrint "Input name: "
'UserIn = CGet()
'If Not UserIn = "" Then
'CPrint "Hi, " & UserIn & "!!" & vbCrLf
'End If
'Form1.Show

Do
 DoEvents
  ReceiveConsole
 DoEvents
Loop
End Sub

Public Sub ReceiveConsole()
Dim Receive As String
Dim I As Integer
Dim CurrendPath As String
Dim TotalFileSize As Variant
Dim ThisFileSize As Long
'Do
DoEvents
    If Right(Form1.File1.Path, 1) <> "\" Then
     CPrint UCase(Form1.Dir1.Path & "\>")
     CurrendPath = Form1.Dir1.Path & "\"
    Else
     CPrint UCase(Form1.Dir1.Path & ">")
     CurrendPath = Form1.Dir1.Path
    End If
DoEvents
    Receive = CGet()
DoEvents
    Select Case LCase(Receive)
     Case "time"
      CPrint Time & vbCrLf
      Exit Sub
     Case "date"
      CPrint Date & vbCrLf
      Exit Sub
     Case "exit"
      FreeConsole
      End
      Unload Form1
      End
     Case "quit"
      FreeConsole
      End
      Unload Form1
      End
     Case "cp"
      CPrint GetConsoleCP
      Exit Sub
     Case "dir"
     TotalFileSize = 0
     CPrint "Directory of " & Form1.Dir1.Path & vbCrLf
     Dim TheSpace As Long
     For I = 0 To Form1.Dir1.ListCount - 1
      TheSpace = Spacing - Len(Form1.Dir1.List(I))
      CPrint Mid(UCase(Form1.Dir1.List(I)), 4) & Space$(TheSpace) & "<DIR>" & vbCrLf
     Next I
     For I = 0 To Form1.File1.ListCount - 1
      ThisFileSize = FileLen(CurrendPath & Form1.File1.List(I))
      TheSpace = Spacing - Len(Form1.File1.List(I))
      CPrint Form1.File1.List(I) & Space$(TheSpace - 3) & ThisFileSize & vbTab & FileDateTime(CurrendPath & Form1.File1.List(I)) & vbCrLf
      TotalFileSize = TotalFileSize + ThisFileSize
     Next I
      CPrint vbTab & vbTab & (Form1.Dir1.ListCount - 1) & " Folder(s)" & vbTab & TotalFileSize & " bytes" & vbCrLf
      Dim FreeSpace As Long
      FreeSpace = GetDriveSpace(Left$(Form1.Drive1.Drive, 2) & "\")
      CPrint vbTab & vbTab & (Form1.File1.ListCount - 1) & " File(s)" & vbTab & FreeSpace & " bytes available" & vbCrLf
      
      'CPrint "Total file size: " & TotalFileSize & vbCrLf
     Exit Sub
     End Select
     
     
     If Left(Receive, 2) = "cd" Then
      If Mid(Receive, 3, 1) = " " Then
       On Error Resume Next
        Form1.Dir1.Path = Mid(Receive, 4)  'ChDir Mid(Receive, 4)
        If Err.Number > 0 Then HandleError Err.Number
      ElseIf Mid(Receive, 3, 1) = "\" Then
        Form1.Dir1.Path = Left(Form1.Drive1.Drive, 2) & "\"
      ElseIf Mid(Receive, 3, 2) = ".." Then
        Dim SlashPos As Long
        Dim SearchChar As String, SearchString As String
        SearchChar = "\"
        SearchString = ReverseString(CurrendPath)
        SlashPos = InStr(1, SearchString, SearchChar)
        SlashPos = InStr(SlashPos + 1, SearchString, SearchChar)
        Form1.Dir1.Path = Left(CurrendPath, (Len(CurrendPath) - SlashPos))
      End If
      
    Form1.Drive1.Refresh
    Form1.Dir1.Refresh
    Form1.File1.Refresh
    Exit Sub
   End If
     
    If Right(LCase(Receive), 1) = ":" Then
     On Error Resume Next
      Form1.Drive1.Drive = Left(LCase(Receive), 2)
      If Err.Number > 0 Then HandleError Err.Number
      Form1.Drive1.Refresh
      Form1.Dir1.Refresh
      Form1.File1.Refresh
     Exit Sub
    End If
    
    On Error Resume Next
    Shell Receive
    If Err.Number <> 0 Then HandleError Err.Number
     
     
'DoEvents
'Loop

End Sub

Public Function GetPath() As String
If Right(Form1.File1.Path, 1) <> "\" Then
 GetPath = Form1.File1.Path & "\"
Else
 GetPath = Form1.File1.Path
End If
End Function

Public Sub HandleError(ErrorNumber As Long)
Debug.Print ErrorNumber
'CPrint "Error: " & Err.Description & vbCrLf
If ErrorNumber = 68 Then
CPrint "Could not access this drive" & vbCrLf
End If

If ErrorNumber = 76 Then
CPrint "Could not access this directory" & vbCrLf
End If

If ErrorNumber = 53 Then
CPrint "Bad file or command" & vbCrLf
End If
End Sub

Public Function ReverseString(Text As String) As String
Dim I As Integer
 Dim NewStr As String
 For I = 1 To Len(Text)
  NewStr = Mid(Text, I, 1) & NewStr
 Next I
ReverseString = NewStr
 
End Function

Public Function GetDriveSpace(Drivez As String) As Long
Dim ReturnVal As Long
Dim Sectors As Long, Bytes As Long, FreeClusters As Long, TotalClusters As Long
 ReturnVal = GetDiskFreeSpace(Drivez, Sectors, Bytes, FreeClusters, TotalClusters)
  GetDriveSpace = (Sectors * Bytes * FreeClusters)
End Function

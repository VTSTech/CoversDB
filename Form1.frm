VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CoversDB (Alpha3)"
   ClientHeight    =   4905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Folder"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Rename"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
   Begin VB.OptionButton Option3 
      Caption         =   "OPL"
      Height          =   195
      Left            =   6120
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "NAME"
      Height          =   195
      Left            =   5280
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ID"
      Height          =   195
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   6615
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   3000
      Width           =   6615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "PS2 - NTSC-U"
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Have: 0"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Console: Not Set"
      Height          =   195
      Left            =   3960
      TabIndex        =   16
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "coversdb.nigeltodman.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4800
      TabIndex        =   15
      Top             =   4680
      Width           =   1905
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Written by Nigel Todman"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   4680
      Width           =   2115
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "File Naming:"
      Height          =   195
      Left            =   3960
      TabIndex        =   13
      Top             =   360
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Not Set"
      Height          =   195
      Left            =   4920
      TabIndex        =   12
      Top             =   360
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Folder: Not Set"
      Height          =   195
      Left            =   2880
      TabIndex        =   11
      Top             =   1080
      Width           =   1065
   End
   Begin VB.Label Label2 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Games: 0"
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y, z, a, ps2_title, ps2_name, ps2_id, fn, tmp, strin, strout, folder, fso, Build
Dim psxdb, ps2db, curr_format, mode, good_cnt, console
Dim nes_name, nes_id, nes_title
Dim game_name, game_id, game_title
Dim console_total() As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Combo1_Click()
If Combo1.Text = "NES - NTSC-U" Then
    console = "nes"
    Label8.Caption = "Console: " & UCase(console)
    a = ListConsole()
ElseIf Combo1.Text = "SNES - NTSC-U" Then
    MsgBox "Not supported yet"
ElseIf Combo1.Text = "GENS - NTSC-U" Then
    MsgBox "Not supported yet"
ElseIf Combo1.Text = "SAT - NTSC-U" Then
    MsgBox "Not supported yet"
ElseIf Combo1.Text = "PSX - NTSC-U" Then
    MsgBox "Not supported yet"
ElseIf Combo1.Text = "GC - NTSC-U" Then
    MsgBox "Not supported yet"
ElseIf Combo1.Text = "WII - NTSC-U" Then
    MsgBox "Not supported yet"
ElseIf Combo1.Text = "PS2 - NTSC-U" Then
    console = "ps2"
    Label8.Caption = "Console: " & UCase(console)
    a = ListConsole()
ElseIf Combo1.Text = "PS3 - NTSC-U" Then
    MsgBox "Not supported yet"
ElseIf Combo1.Text = "XBOX - NTSC-U" Then
    MsgBox "Not supported yet"
ElseIf Combo1.Text = "X360 - NTSC-U" Then
    MsgBox "Not supported yet"
ElseIf Combo1.Text = "PS2 - PAL" Then
    console = "ps2"
    Label8.Caption = "Console: " & UCase(console)
    a = ListConsole()
Else
    MsgBox "Not supported yet"
End If
End Sub

Private Sub Command1_Click()
If Option1.Value = True Then
    mode = "id"
ElseIf Option2.Value = True Then
    mode = "name"
ElseIf Option3.Value = True Then
    mode = "opl"
End If
If mode = "" Then
    MsgBox "Error: Select a Filename first"
Else
    a = CheckConsole(mode)
End If
End Sub

Private Function RenameConsole(mode)
If curr_format = "name" Then
    If mode = "id" Then
        good_cnt = 0
        strout = ""
        For z = 0 To UBound(console_total) - 1
            tmp = Split(console_total(z), ";")
            game_id = tmp(0)
            game_title = tmp(1)
            game_name = ImgFN(tmp(1)) & ".jpg"
            If console = "ps2" Then
                ps2_opl = PS2toOPL(tmp(0))
            End If
            If fso.FileExists(folder & game_name) Then
                good_cnt = good_cnt + 1
                'MsgBox "cmd.exe /c " & Chr(34) & "ren " & folder & ps2_name & " " & ps2_id & ".jpg" & Chr(34)
                strout = strout & "ren " & Chr(34) & folder & game_name & Chr(34) & " " & Replace(game_id, " ", "") & ".jpg" & vbCrLf
                Sleep (10)
            End If
            If console = "ps2" Then
                Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & ps2_opl
            Else
                Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & Replace(game_id, " ", "") & ".jpg"
            End If
        Next z
        Close #2
        Open VB.App.Path & "\tmp.cmd" For Output As #2
        Print #2, strout
        Sleep (10)
        Close #2
        a = MsgBox("Execute Rename Script generated at " & VB.App.Path & "\tmp.cmd", vbYesNo)
        If a = vbYes Then
            Shell (VB.App.Path & "\tmp.cmd")
        End If
    ElseIf mode = "opl" Then
        good_cnt = 0
        strout = ""
        For z = 0 To UBound(console_total) - 1
            tmp = Split(console_total(z), ";")
            game_id = tmp(0)
            game_title = tmp(1)
            game_name = ImgFN(tmp(1)) & ".jpg"
            If console = "ps2" Then
                ps2_opl = PS2toOPL(tmp(0))
            End If
            If fso.FileExists(folder & game_name) Then
                good_cnt = good_cnt + 1
                'MsgBox "cmd.exe /c " & Chr(34) & "ren " & folder & ps2_name & " " & ps2_id & ".jpg" & Chr(34)
                strout = strout & "ren " & Chr(34) & folder & ps2_name & " " & ps2_opl & Chr(34) & vbCrLf
                Sleep (10)
            End If
            If console = "ps2" Then
                Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & ps2_opl
            Else
                Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & Replace(game_id, " ", "") & ".jpg"
            End If
        Next z
        Close #2
        Open VB.App.Path & "\tmp.cmd" For Output As #2
        Print #2, strout
        Sleep (10)
        Close #2
        a = MsgBox("Execute Rename Script generated at " & VB.App.Path & "\tmp.cmd", vbYesNo)
        If a = vbYes Then
            Shell (VB.App.Path & "\tmp.cmd")
        End If
    End If
ElseIf curr_format = "id" Then
    If mode = "name" Then
        good_cnt = 0
        strout = ""
        For z = 0 To UBound(console_total) - 1
            tmp = Split(console_total(z), ";")
            game_id = tmp(0)
            game_title = tmp(1)
            game_name = ImgFN(tmp(1)) & ".jpg"
            ps2_opl = PS2toOPL(tmp(0))
            If fso.FileExists(folder & Replace(game_id, " ", "") & ".jpg") Then
                good_cnt = good_cnt + 1
                'MsgBox "cmd.exe /c " & Chr(34) & "ren " & folder & ps2_name & " " & ps2_id & ".jpg" & Chr(34)
                strout = strout & "ren " & Chr(34) & folder & game_id & ".jpg" & Chr(34) & " " & game_name & vbCrLf
                Sleep (10)
            End If
            If console = "ps2" Then
                Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & ps2_opl
            Else
                Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & Replace(game_id, " ", "") & ".jpg"
            End If
        Next z
        Close #2
        Open VB.App.Path & "\tmp.cmd" For Output As #2
        Print #2, strout
        Sleep (10)
        Close #2
        a = MsgBox("Execute Rename Script generated at " & VB.App.Path & "\tmp.cmd", vbYesNo)
        If a = vbYes Then
            Shell (VB.App.Path & "\tmp.cmd")
        End If
    ElseIf mode = "opl" Then
        good_cnt = 0
        strout = ""
        For z = 0 To UBound(console_total) - 1
            tmp = Split(console_total(z), ";")
            game_id = tmp(0)
            game_title = tmp(1)
            game_name = ImgFN(tmp(1)) & ".jpg"
            ps2_opl = PS2toOPL(tmp(0))
            If fso.FileExists(folder & ps2_id & ".jpg") Then
                good_cnt = good_cnt + 1
                'MsgBox "cmd.exe /c " & Chr(34) & "ren " & folder & ps2_name & " " & ps2_id & ".jpg" & Chr(34)
                strout = strout & "ren " & Chr(34) & folder & ps2_id & ".jpg" & Chr(34) & " " & ps2_opl & vbCrLf
                Sleep (10)
            End If
            If console = "ps2" Then
                Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & ps2_opl
            Else
                Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & Replace(game_id, " ", "") & ".jpg"
            End If
        Next z
        Close #2
        Open VB.App.Path & "\tmp.cmd" For Output As #2
        Print #2, strout
        Sleep (10)
        Close #2
        a = MsgBox("Execute Rename Script generated at " & VB.App.Path & "\tmp.cmd", vbYesNo)
        If a = vbYes Then
            Shell (VB.App.Path & "\tmp.cmd")
        End If
    End If
ElseIf curr_format = "opl" Then
    If mode = "name" Then
        good_cnt = 0
        strout = ""
        For z = 0 To UBound(console_total) - 1
            tmp = Split(console_total(z), ";")
            game_id = tmp(0)
            game_title = tmp(1)
            game_name = ImgFN(tmp(1)) & ".jpg"
            ps2_opl = PS2toOPL(tmp(0))
            If fso.FileExists(folder & ps2_opl) Then
                good_cnt = good_cnt + 1
                'MsgBox "cmd.exe /c " & Chr(34) & "ren " & folder & ps2_name & " " & ps2_id & ".jpg" & Chr(34)
                strout = strout & "ren " & folder & ps2_opl & " " & ps2_name & vbCrLf
                Sleep (10)
            End If
            If console = "ps2" Then
                Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & ps2_opl
            Else
                Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & Replace(game_id, " ", "") & ".jpg"
            End If
        Next z
        Close #2
        Open VB.App.Path & "\tmp.cmd" For Output As #2
        Print #2, strout
        Sleep (10)
        Close #2
        a = MsgBox("Execute Rename Script generated at " & VB.App.Path & "\tmp.cmd", vbYesNo)
        If a = vbYes Then
            Shell (VB.App.Path & "\tmp.cmd")
        End If
    ElseIf mode = "id" Then
        good_cnt = 0
        strout = ""
        For z = 0 To UBound(console_total) - 1
            tmp = Split(console_total(z), ";")
            game_id = tmp(0)
            game_title = tmp(1)
            game_name = ImgFN(tmp(1)) & ".jpg"
            ps2_opl = PS2toOPL(tmp(0))
            If fso.FileExists(folder & ps2_opl) Then
                good_cnt = good_cnt + 1
                'MsgBox "cmd.exe /c " & Chr(34) & "ren " & folder & ps2_name & " " & ps2_id & ".jpg" & Chr(34)
                strout = strout & "ren " & folder & ps2_opl & " " & ps2_id & ".jpg" & vbCrLf
                Sleep (10)
            End If
            If console = "ps2" Then
                Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & ps2_opl
            Else
                Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & Replace(game_id, " ", "") & ".jpg"
            End If
        Next z
        Close #2
        Open VB.App.Path & "\tmp.cmd" For Output As #2
        Print #2, strout
        Sleep (10)
        Close #2
        a = MsgBox("Execute Rename Script generated at " & VB.App.Path & "\tmp.cmd", vbYesNo)
        If a = vbYes Then
            Shell (VB.App.Path & "\tmp.cmd")
        End If
    End If
End If
End Function

Private Sub Command2_Click()
If curr_format = "Not Set" Then
MsgBox "Check first to set File Naming"
Else
    If Option1.Value = True Then
        RenameConsole ("id")
    ElseIf Option2.Value = True Then
        RenameConsole ("name")
    ElseIf Option3.Value = True Then
        RenameConsole ("opl")
    End If
End If
End Sub

Private Sub Command3_Click()
Label3.Caption = InputBox("Enter Folder Path:")
End Sub
Private Function CheckConsole(mode)
'MsgBox List1.ListCount
If mode = "name" Then
    good_cnt = 0
    For z = 0 To UBound(console_total) - 1
        tmp = Split(console_total(z), ";")
        game_id = tmp(0)
        game_title = tmp(1)
        game_name = ImgFN(tmp(1)) & ".jpg"
        ps2_opl = PS2toOPL(tmp(0))
        If fso.FileExists(folder & game_name) Then
            good_cnt = good_cnt + 1
        End If
        If console = "ps2" Then
            Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & ps2_opl
        Else
            Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & Replace(game_id, " ", "") & ".jpg"
        End If
    Next z
    If good_cnt >= 1 Then
        curr_format = "name"
        Label9.Caption = "Have: " & good_cnt
    End If
    Label4.Caption = curr_format
ElseIf mode = "id" Then
    good_cnt = 0
    For z = 0 To UBound(console_total) - 1
        tmp = Split(console_total(z), ";")
        game_id = tmp(0)
        game_title = tmp(1)
        game_name = ImgFN(tmp(1)) & ".jpg"
        ps2_opl = PS2toOPL(tmp(0))
        If fso.FileExists(folder & game_id & ".jpg") Then
            good_cnt = good_cnt + 1
        End If
        If console = "ps2" Then
            Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & ps2_opl
        Else
            Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & Replace(game_id, " ", "") & ".jpg"
        End If
    Next z
    If good_cnt >= 1 Then
        curr_format = "id"
        Label9.Caption = "Have: " & good_cnt
    End If
    Label4.Caption = curr_format
ElseIf mode = "opl" Then
    good_cnt = 0
    For z = 0 To UBound(console_total) - 1
        tmp = Split(console_total(z), ";")
        game_id = tmp(0)
        game_title = tmp(1)
        game_name = ImgFN(tmp(1)) & ".jpg"
        ps2_opl = PS2toOPL(tmp(0))
        If fso.FileExists(folder & ps2_opl) Then
            good_cnt = good_cnt + 1
        End If
        If console = "ps2" Then
            Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & ps2_opl
        Else
            Text1.Text = game_id & vbCrLf & game_title & vbCrLf & game_name & vbCrLf & Replace(game_id, " ", "") & ".jpg"
        End If
    Next z
    If good_cnt >= 1 Then
        curr_format = "opl"
        Label9.Caption = "Have: " & good_cnt
    End If
    Label4.Caption = curr_format
End If
End Function
Public Function ListConsole()
If console = "ps2" Then
    If Combo1.Text = "PS2 - NTSC-U" Then
        CoversDB = VB.App.Path & "\dat\PS2_NTSCU.dat"
        folder = VB.App.Path & "\covers\PS2\"
    ElseIf Combo1.Text = "PS2 - PAL" Then
        CoversDB = VB.App.Path & "\dat\PS2_PAL.dat"
        folder = VB.App.Path & "\covers\PS2\"
    End If
    Label3.Caption = folder
ElseIf console = "nes" Then
    If Combo1.Text = "NES - NTSC-U" Then
        CoversDB = VB.App.Path & "\dat\NES_NTSCU.dat"
        folder = VB.App.Path & "\covers\NES\"
    ElseIf Combo1.Text = "NES - PAL" Then
        CoversDB = VB.App.Path & "\dat\NES_PAL.dat"
        folder = VB.App.Path & "\covers\NES\"
    End If
    Label3.Caption = folder
End If
If fso.FileExists(CoversDB) Then
    fn = CoversDB
    x = 0
    List1.Clear
    Close #1
    Open fn For Input As #1
    Do Until EOF(1)
        Line Input #1, tmp
        x = x + 1
        tmp = tmp & tmp & vbCrLf
    Loop
    Close #1
    ReDim console_total(x)
    Label1.Caption = "Total: " & x
    x = 0
    Close #1
    Open fn For Input As #1
    Do Until EOF(1)
        Line Input #1, console_total(x)
        x = x + 1
    Loop
    Close #1
    For y = 0 To UBound(console_total)
        List1.AddItem console_total(y)
    Next y
Else
    MsgBox "Error: CoversDB for this console does not exist in \dat\"
End If
End Function
Private Sub Form_Load()
Set fso = CreateObject("Scripting.FileSystemObject")
Build = "0.0.1-ALPHA4"
Form1.Caption = "CoversDB v" & Build
Text1.Text = ""
folder = "Not Set"
curr_format = "Not Set"
console = "Not Set"
Label3.Caption = folder
Label4.Caption = curr_format
x = 0
y = 0
z = 0
Combo1.AddItem "NES - NTSC-U"
Combo1.AddItem "SNES - NTSC-U"
Combo1.AddItem "GENS - NTSC-U"
Combo1.AddItem "SAT - NTSC-U"
Combo1.AddItem "PSX - NTSC-U"
Combo1.AddItem "GC - NTSC-U"
Combo1.AddItem "WII - NTSC-U"
Combo1.AddItem "PS2 - NTSC-U"
Combo1.AddItem "PS3 - NTSC-U"
Combo1.AddItem "XBOX - NTSC-U"
Combo1.AddItem "X360 - NTSC-U"
Combo1.AddItem "NES - PAL"
Combo1.AddItem "SNES - PAL"
Combo1.AddItem "GENS - PAL"
Combo1.AddItem "SAT - PAL"
Combo1.AddItem "PSX - PAL"
Combo1.AddItem "GC - PAL"
Combo1.AddItem "WII - PAL"
Combo1.AddItem "PS2 - PAL"
Combo1.AddItem "PS3 - PAL"
Combo1.AddItem "XBOX - PAL"
Combo1.AddItem "X360 - PAL"
Combo1.Text = "Select console..."
End Sub
Private Function ImgFN(strin)
strin = Replace(strin, " (USA)", "")
strin = Replace(strin, " (Disc 1)", "")
strin = Replace(strin, " (Disc 2)", "")
strin = Replace(strin, " (Disc 3)", "")
strin = Replace(strin, " (Disc 4)", "")
strin = Replace(strin, " (Greatest Hits)", "")
strin = Replace(strin, " (En,Fr)", "")
strin = Replace(strin, " (En,Es)", "")
strin = Replace(strin, " (En,Ja)", "")
strin = Replace(strin, " (En,Fr,De)", "")
strin = Replace(strin, " (En,Fr,Es)", "")
strin = Replace(strin, " (En,Fr,De,Es)", "")
strin = Replace(strin, " (En,Fr,De,It)", "")
strin = Replace(strin, " (En,Fr,De,Es,It)", "")
strin = Replace(strin, " (En,De,Es,Nl,Sv)", "")
strin = Replace(strin, " (En,Ja,Fr,De,Es,It)", "")
strin = Replace(strin, " (En,Fr,De,Es,It,Pt,Ru)", "")
strin = Replace(strin, " (En,Ja,Fr,De,Es,It,Ko)", "")
strin = Replace(strin, " (En,Fr,De,Es,It,Nl,Sv,Da)", "")
strin = Replace(strin, " (JU)", "")
strin = Replace(strin, ".zip", "")
strin = Replace(strin, ".7z", "")
strin = Replace(strin, " - ", "_")
strin = Replace(strin, " ", "_")
strin = Replace(strin, "-", "_")
strin = Replace(strin, "'", "")
strin = Replace(strin, ",", "")
strin = Replace(strin, "vol.", "vol_")
strin = Replace(strin, "#", "_")
strin = Replace(strin, ".", "")
strin = Replace(strin, "[", "")
strin = Replace(strin, "]", "")
strin = LCase(strin)
ImgFN = strin
End Function
Private Function PS2toOPL(strin)
strin = Mid(strin, 1, 4) & "_" & Mid(strin, 6, 3) & "." & Mid(strin, 8, 2) & "_COV.jpg"
PS2toOPL = strin
End Function

Private Sub Label7_Click()
Shell ("cmd.exe /c start https://coversdb.nigeltodman.com"), vbHide
End Sub

Private Sub List1_Click()
tmp = Split(List1.List(List1.ListIndex), ";")
If console = "ps2" Then
    ps2_id = tmp(0)
    ps2_title = tmp(1)
    ps2_name = ImgFN(tmp(1)) & ".jpg"
    ps2_opl = PS2toOPL(tmp(0))
    Text1.Text = ps2_id & vbCrLf & ps2_title & vbCrLf & ps2_name & vbCrLf & ps2_opl
ElseIf console = "nes" Then
    nes_id = tmp(0)
    nes_title = tmp(1)
    nes_name = ImgFN(tmp(1)) & ".jpg"
    'nes_opl = PS2toOPL(tmp(0))
    Text1.Text = nes_id & vbCrLf & nes_title & vbCrLf & nes_name & vbCrLf & nes_id & ".jpg" & vbCrLf
End If
End Sub

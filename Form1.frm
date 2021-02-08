VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Text            =   "56000"
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Text            =   "0"
      Top             =   7560
      Width           =   5775
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   7200
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1345
      TabIndex        =   6
      Top             =   4080
      Width           =   4425
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   11
      Top             =   4560
      Width           =   5775
   End
   Begin VB.Label Label4 
      Caption         =   "Average deals:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "One draw Royal Flush Odds:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Cards dealt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Attempts: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean

Private Sub Command1_Click()

If Text2 = "" Then: Text2 = 1

If IsNumeric(Text2) = False Then

    MsgBox "Please use numeric values only for the timeout delay", vbExclamation, "card odds"
    
    Exit Sub
    
End If

If Text2 < 0 Then

    MsgBox "Please use a positive number for the timeout delay", vbExclamation, "card odds"

    Exit Sub
    
End If

Label1 = "Attempts: 0"

timeout 0.1

flag = False

Do

If flag Then
Exit Sub
End If

If timeout_delay = Text2 Then
DoEvents
timeout_delay = 0
End If

timeout_delay = timeout_delay + 1

attempts = attempts + 1

Label1 = "Attempts: " & attempts

try_again:

buffer1 = ""
buffer2 = ""
buffer3 = ""
buffer4 = ""
buffer5 = ""

Randomize

card1 = Int((52 * Rnd) + 1)
card2 = Int((52 * Rnd) + 1)
card3 = Int((52 * Rnd) + 1)
card4 = Int((52 * Rnd) + 1)
card5 = Int((52 * Rnd) + 1)

If card1 = card2 Then
GoTo try_again:
ElseIf card1 = card3 Then
GoTo try_again:
ElseIf card1 = card4 Then
GoTo try_again:
ElseIf card1 = card5 Then
GoTo try_again:
End If

If card2 = card1 Then
GoTo try_again:
ElseIf card2 = card3 Then
GoTo try_again:
ElseIf card2 = card4 Then
GoTo try_again:
ElseIf card2 = card5 Then
GoTo try_again:
End If

If card3 = card1 Then
GoTo try_again:
ElseIf card3 = card2 Then
GoTo try_again:
ElseIf card3 = card4 Then
GoTo try_again:
ElseIf card3 = card5 Then
GoTo try_again:
End If

If card4 = card1 Then
GoTo try_again:
ElseIf card4 = card2 Then
GoTo try_again:
ElseIf card4 = card3 Then
GoTo try_again:
ElseIf card4 = card5 Then
GoTo try_again:
End If

If card5 = card1 Then
GoTo try_again:
ElseIf card5 = card2 Then
GoTo try_again:
ElseIf card5 = card3 Then
GoTo try_again:
ElseIf card5 = card4 Then
GoTo try_again:
End If

Select Case card1
Case 9
buffer1 = "10"
Case 10
buffer1 = "J"
Case 11
buffer1 = "Q"
Case 12
buffer1 = "K"
Case 13
buffer1 = "A"
Case 22
buffer1 = "10"
Case 23
buffer1 = "J"
Case 24
buffer1 = "Q"
Case 25
buffer1 = "K"
Case 26
buffer1 = "A"
Case 35
buffer1 = "10"
Case 36
buffer1 = "J"
Case 37
buffer1 = "Q"
Case 38
buffer1 = "K"
Case 39
buffer1 = "A"
Case 48
buffer1 = "10"
Case 49
buffer1 = "J"
Case 50
buffer1 = "Q"
Case 51
buffer1 = "K"
Case 52
buffer1 = "A"
End Select

Select Case card2
Case 9
buffer2 = "10"
Case 10
buffer2 = "J"
Case 11
buffer2 = "Q"
Case 12
buffer2 = "K"
Case 13
buffer2 = "A"
Case 22
buffer2 = "10"
Case 23
buffer2 = "J"
Case 24
buffer2 = "Q"
Case 25
buffer2 = "K"
Case 26
buffer2 = "A"
Case 35
buffer2 = "10"
Case 36
buffer2 = "J"
Case 37
buffer2 = "Q"
Case 38
buffer2 = "K"
Case 39
buffer2 = "A"
Case 48
buffer2 = "10"
Case 49
buffer2 = "J"
Case 50
buffer2 = "Q"
Case 51
buffer2 = "K"
Case 52
buffer2 = "A"
End Select

Select Case card3
Case 9
buffer3 = "10"
Case 10
buffer3 = "J"
Case 11
buffer3 = "Q"
Case 12
buffer3 = "K"
Case 13
buffer3 = "A"
Case 22
buffer3 = "10"
Case 23
buffer3 = "J"
Case 24
buffer3 = "Q"
Case 25
buffer3 = "K"
Case 26
buffer3 = "A"
Case 35
buffer3 = "10"
Case 36
buffer3 = "J"
Case 37
buffer3 = "Q"
Case 38
buffer3 = "K"
Case 39
buffer3 = "A"
Case 48
buffer3 = "10"
Case 49
buffer3 = "J"
Case 50
buffer3 = "Q"
Case 51
buffer3 = "K"
Case 52
buffer3 = "A"
End Select

Select Case card4
Case 9
buffer4 = "10"
Case 10
buffer4 = "J"
Case 11
buffer4 = "Q"
Case 12
buffer4 = "K"
Case 13
buffer4 = "A"
Case 22
buffer4 = "10"
Case 23
buffer4 = "J"
Case 24
buffer4 = "Q"
Case 25
buffer4 = "K"
Case 26
buffer4 = "A"
Case 35
buffer4 = "10"
Case 36
buffer4 = "J"
Case 37
buffer4 = "Q"
Case 38
buffer4 = "K"
Case 39
buffer4 = "A"
Case 48
buffer4 = "10"
Case 49
buffer4 = "J"
Case 50
buffer4 = "Q"
Case 51
buffer4 = "K"
Case 52
buffer4 = "A"
End Select

Select Case card5
Case 9
buffer5 = "10"
Case 10
buffer5 = "J"
Case 11
buffer5 = "Q"
Case 12
buffer5 = "K"
Case 13
buffer5 = "A"
Case 22
buffer5 = "10"
Case 23
buffer5 = "J"
Case 24
buffer5 = "Q"
Case 25
buffer5 = "K"
Case 26
buffer5 = "A"
Case 35
buffer5 = "10"
Case 36
buffer5 = "J"
Case 37
buffer5 = "Q"
Case 38
buffer5 = "K"
Case 39
buffer5 = "A"
Case 48
buffer5 = "10"
Case 49
buffer5 = "J"
Case 50
buffer5 = "Q"
Case 51
buffer5 = "K"
Case 52
buffer5 = "A"
End Select

If buffer1 <> "" And buffer2 <> "" And buffer3 <> "" And buffer4 <> "" And buffer5 <> "" Then

    If card1 >= 9 And card1 <= 13 Then
    If card2 >= 9 And card2 <= 13 Then
    If card3 >= 9 And card3 <= 13 Then
    If card4 >= 9 And card4 <= 13 Then
    If card5 >= 9 And card5 <= 13 Then

    List2.AddItem "Spade Royal Flush in: " & attempts & " card deals."

    List1.AddItem buffer1 & "," & buffer2 & "," & buffer3 & "," & buffer4 & "," & buffer5

    List3.AddItem attempts

    Text1.Text = Text1.Text + attempts
    
    average = Text1.Text \ List2.ListCount
    
    Label4 = "Average deals: " & average

    attempts = 0

    completed = completed + 1

    buffer1 = ""
    buffer2 = ""
    buffer3 = ""
    buffer4 = ""
    buffer5 = ""
    card1 = ""
    card2 = ""
    card3 = ""
    card4 = ""
    card5 = ""
    timeout 0.1

    If completed = 200 Then
    Exit Sub
    End If

    GoTo bypass:
    
    End If
    End If
    End If
    End If
    End If

    If card1 >= 22 And card1 <= 26 Then
    If card2 >= 22 And card2 <= 26 Then
    If card3 >= 22 And card3 <= 26 Then
    If card4 >= 22 And card4 <= 26 Then
    If card5 >= 22 And card5 <= 26 Then

    List2.AddItem "Diamond Royal Flush in: " & attempts & " card deals."

    List1.AddItem buffer1 & "," & buffer2 & "," & buffer3 & "," & buffer4 & "," & buffer5

    List3.AddItem attempts

    Text1.Text = Text1.Text + attempts
    
    average = Text1.Text \ List2.ListCount
    
    Label4 = "Average deals: " & average

    attempts = 0

    completed = completed + 1

    buffer1 = ""
    buffer2 = ""
    buffer3 = ""
    buffer4 = ""
    buffer5 = ""
    card1 = ""
    card2 = ""
    card3 = ""
    card4 = ""
    card5 = ""
    timeout 0.1

    If completed = 200 Then
    Exit Sub
    End If

    GoTo bypass:

    End If
    End If
    End If
    End If
    End If
    
    If card1 >= 35 And card1 <= 39 Then
    If card2 >= 35 And card2 <= 39 Then
    If card3 >= 35 And card3 <= 39 Then
    If card4 >= 35 And card4 <= 39 Then
    If card5 >= 35 And card5 <= 39 Then

    List2.AddItem "Club Royal Flush in: " & attempts & " card deals."

    List1.AddItem buffer1 & "," & buffer2 & "," & buffer3 & "," & buffer4 & "," & buffer5

    List3.AddItem attempts

    Text1.Text = Text1.Text + attempts
    
    average = Text1.Text \ List2.ListCount
    
    Label4 = "Average deals: " & average

    attempts = 0

    completed = completed + 1

    buffer1 = ""
    buffer2 = ""
    buffer3 = ""
    buffer4 = ""
    buffer5 = ""
    card1 = ""
    card2 = ""
    card3 = ""
    card4 = ""
    card5 = ""
    timeout 0.1

    If completed = 200 Then
    Exit Sub
    End If

    GoTo bypass:

    End If
    End If
    End If
    End If
    End If
    
    If card1 >= 48 And card1 <= 52 Then
    If card2 >= 48 And card2 <= 52 Then
    If card3 >= 48 And card3 <= 52 Then
    If card4 >= 48 And card4 <= 52 Then
    If card5 >= 48 And card5 <= 52 Then

    List2.AddItem "Heart Royal Flush in: " & attempts & " card deals."

    List1.AddItem buffer1 & "," & buffer2 & "," & buffer3 & "," & buffer4 & "," & buffer5

    List3.AddItem attempts

    Text1.Text = Text1.Text + attempts
    
    average = Text1.Text \ List2.ListCount
    
    Label4 = "Average deals: " & average

    attempts = 0

    completed = completed + 1

    buffer1 = ""
    buffer2 = ""
    buffer3 = ""
    buffer4 = ""
    buffer5 = ""
    card1 = ""
    card2 = ""
    card3 = ""
    card4 = ""
    card5 = ""
    timeout 0.1

    If completed = 200 Then
    Exit Sub
    End If

    GoTo bypass:

    End If
    End If
    End If
    End If
    End If
    
End If

bypass:

Loop

End Sub


Sub Command3_Click()


End Sub

Private Sub Command2_Click()
flag = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End
End Sub

Private Sub List2_Click()
a = List2.ListIndex
List1.ListIndex = a
End Sub


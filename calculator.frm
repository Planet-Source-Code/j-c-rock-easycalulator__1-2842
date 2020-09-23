VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Easy Calculator"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "calculator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSin 
      Caption         =   "Sin"
      Height          =   540
      Left            =   3675
      TabIndex        =   30
      Top             =   1995
      Width           =   540
   End
   Begin VB.CommandButton cmdCos 
      Caption         =   "Cos"
      Height          =   540
      Left            =   3675
      TabIndex        =   29
      Top             =   1365
      Width           =   540
   End
   Begin VB.CommandButton cmdSquareRoot 
      Caption         =   "sqrt"
      Height          =   540
      Left            =   3675
      TabIndex        =   28
      Top             =   735
      Width           =   540
   End
   Begin VB.Frame Frame2 
      Caption         =   "Memory Contents"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2730
      TabIndex        =   26
      Top             =   3255
      Width           =   2535
      Begin VB.Label lblMemory1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FF00&
         Height          =   330
         Left            =   105
         TabIndex        =   27
         Top             =   210
         Width           =   2325
      End
   End
   Begin VB.CommandButton cmdMemory3 
      Caption         =   "MC"
      Height          =   435
      Left            =   4830
      TabIndex        =   25
      Top             =   1680
      Width           =   540
   End
   Begin VB.CommandButton cmdMemory2 
      Caption         =   "MR"
      Height          =   435
      Left            =   4830
      TabIndex        =   24
      Top             =   2205
      Width           =   540
   End
   Begin VB.CommandButton cmdMemory1 
      Caption         =   "M+"
      Height          =   435
      Left            =   4830
      TabIndex        =   23
      Top             =   2730
      Width           =   540
   End
   Begin VB.Frame Frame1 
      Caption         =   "Running Total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   98
      TabIndex        =   21
      Top             =   3255
      Width           =   2535
      Begin VB.Label lblRunning 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FF00&
         Height          =   330
         Left            =   105
         TabIndex        =   22
         Top             =   210
         Width           =   2325
      End
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "CA"
      Height          =   435
      Left            =   4830
      TabIndex        =   20
      Top             =   735
      Width           =   540
   End
   Begin VB.CommandButton cmdEquals 
      Caption         =   "="
      Height          =   540
      Left            =   2415
      TabIndex        =   19
      Top             =   2625
      Width           =   1170
   End
   Begin VB.CommandButton cmdDivide 
      Caption         =   "/"
      Height          =   540
      Left            =   3045
      TabIndex        =   18
      Top             =   1995
      Width           =   540
   End
   Begin VB.CommandButton cmdTimes 
      Caption         =   "*"
      Height          =   540
      Left            =   2415
      TabIndex        =   17
      Top             =   1995
      Width           =   540
   End
   Begin VB.CommandButton cmdMinus 
      Caption         =   "-"
      Height          =   540
      Left            =   3045
      TabIndex        =   16
      Top             =   1365
      Width           =   540
   End
   Begin VB.CommandButton cmdPlusMinus 
      Caption         =   "+/-"
      Height          =   540
      Left            =   2415
      TabIndex        =   15
      Top             =   735
      Width           =   540
   End
   Begin VB.CommandButton cmdOver 
      Caption         =   "1/X"
      Height          =   540
      Left            =   3045
      TabIndex        =   14
      Top             =   735
      Width           =   540
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      Height          =   540
      Left            =   2415
      TabIndex        =   13
      Top             =   1365
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1470
      TabIndex        =   12
      Top             =   2625
      Width           =   540
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CE"
      Height          =   435
      Left            =   4830
      TabIndex        =   11
      Top             =   210
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "9"
      Height          =   540
      Index           =   9
      Left            =   1470
      TabIndex        =   10
      Top             =   1995
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "8"
      Height          =   540
      Index           =   8
      Left            =   840
      TabIndex        =   9
      Top             =   1995
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "7"
      Height          =   540
      Index           =   7
      Left            =   210
      TabIndex        =   8
      Top             =   1995
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "6"
      Height          =   540
      Index           =   6
      Left            =   1470
      TabIndex        =   7
      Top             =   1365
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "5"
      Height          =   540
      Index           =   5
      Left            =   840
      TabIndex        =   6
      Top             =   1365
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "4"
      Height          =   540
      Index           =   4
      Left            =   210
      TabIndex        =   5
      Top             =   1365
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "3"
      Height          =   540
      Index           =   3
      Left            =   1470
      TabIndex        =   4
      Top             =   735
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "2"
      Height          =   540
      Index           =   2
      Left            =   840
      TabIndex        =   3
      Top             =   735
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "1"
      Height          =   540
      Index           =   1
      Left            =   210
      TabIndex        =   2
      Top             =   735
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "0"
      Height          =   540
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   2625
      Width           =   540
   End
   Begin VB.Label lblDisplay 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   165
      TabIndex        =   0
      Top             =   210
      Width           =   4590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim operand1 As Double, operand2 As Double
Dim operator As String
Dim cleardisplay As Boolean
Dim memory1 As Double
Dim percent As Double



Private Sub cmdClear_Click()

lblDisplay.Caption = ""

End Sub

Private Sub cmdClearAll_Click()

operand1 = 0
operand2 = 0
lblDisplay.Caption = ""
lblRunning.Caption = ""

End Sub

Private Sub cmdCos_Click()
lblDisplay.Caption = Cos(Val(lblDisplay.Caption))

End Sub

Private Sub cmdDigits_Click(Index As Integer)

If cleardisplay Then
    lblDisplay.Caption = ""
    cleardisplay = False
End If
If Len(lblDisplay.Caption) < 10 Then
   lblDisplay.Caption = lblDisplay.Caption + cmdDigits(Index).Caption
    Else
    End If


End Sub

Private Sub cmdDivide_Click()

operand1 = Val(lblDisplay.Caption)
operator = "/"
lblDisplay.Caption = ""

End Sub

Private Sub cmdEquals_Click()

On Error GoTo errorhandler

Dim result As Double

operand2 = Val(lblDisplay.Caption)
If operator = "+" Then result = operand1 + operand2
If operator = "-" Then result = operand1 - operand2
If operator = "*" Then result = operand1 * operand2
If operator = "/" And operand2 <> "0" Then _
                result = operand1 / operand2
lblDisplay.Caption = result
operand1 = result
lblRunning.Caption = result
Exit Sub

errorhandler:
MsgBox "The operation resulted in the following error" & _
    vbCrLf & Err.Description
lblDisplay.Caption = "ERROR"
cleardisplay = True

End Sub

Private Sub cmdMemory1_Click()
memory1 = lblDisplay.Caption
lblMemory1 = memory1
End Sub

Private Sub cmdMemory2_Click()
lblDisplay.Caption = memory1
End Sub

Private Sub cmdMemory3_Click()
memory1 = 0
lblMemory1.Caption = ""

End Sub

Private Sub cmdMinus_Click()

operand1 = Val(lblDisplay.Caption)
operator = "-"
lblDisplay.Caption = ""

End Sub

Private Sub cmdOver_Click()

If Val(lblDisplay.Caption) <> 0 Then lblDisplay.Caption = _
                        1 / Val(lblDisplay.Caption)
                        
End Sub


Private Sub cmdPlus_Click()

operand1 = Val(lblDisplay.Caption)
operator = "+"
lblDisplay.Caption = ""
lblRunning.Caption = operand1

End Sub

Private Sub cmdPlusMinus_Click()

lblDisplay.Caption = -Val(lblDisplay.Caption)

End Sub

Private Sub cmdSin_Click()
lblDisplay.Caption = Sin(Val(lblDisplay.Caption))
End Sub

Private Sub cmdSquareRoot_Click()
If lblDisplay.Caption < 0 Then
MsgBox "Can't calculate the square root of a negative number"
Else
lblDisplay.Caption = Sqr(Val(lblDisplay.Caption))
End If
End Sub

Private Sub cmdTimes_Click()

operand1 = Val(lblDisplay.Caption)
operator = "*"
lblDisplay.Caption = ""

End Sub

Private Sub Command1_Click()

If InStr(lblDisplay.Caption, ".") Then
    Exit Sub
Else
    lblDisplay.Caption = lblDisplay.Caption + "."
End If

End Sub

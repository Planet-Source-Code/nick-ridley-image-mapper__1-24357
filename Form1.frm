VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00735955&
   Caption         =   "Image Mapper"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00735955&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00735955&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00735955&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B5796F&
      Height          =   285
      Left            =   5280
      TabIndex        =   8
      Text            =   "35"
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00735955&
      Caption         =   "Blend Pictures (mix)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      MaskColor       =   &H00B5796F&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00735955&
      Caption         =   "Extract differences (from 1 compared to 2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      MaskColor       =   &H00B5796F&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00735955&
      Caption         =   "Compare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      MaskColor       =   &H00B5796F&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   3135
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   3480
      Width           =   3255
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00B5796F&
      Height          =   3255
      Left            =   3480
      ScaleHeight     =   3195
      ScaleWidth      =   3195
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00B5796F&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2001 SpyderNet Productions "
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   6600
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tolerence:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00735955&
      BackStyle       =   0  'Transparent
      Caption         =   "The screen resoloution must be set to 800x600 otherwise image mapping/conversion will not work"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   3480
      TabIndex        =   3
      Top             =   3960
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x As Long, y As Long
Dim c As Long, d As Long
Dim p1 As Long, p2 As Long
Dim R1 As Long, B1 As Long, G1 As Long
Dim R2 As Long, B2 As Long, G2 As Long
Dim t As Long, r As Integer
Dim per As Long, msg As String
Picture3.Cls

t = Text1.Text

Form1.Caption = "Image Mapper - Mapping"

Do Until y > Picture1.Height Or y > Picture2.Height

p1 = Picture1.Point(x, y)
p2 = Picture2.Point(x, y)

RGB_get p1, R1, B1, G1
RGB_get p2, R2, B2, G2

Diff r, R1, R2, B1, B2, G1, G2, t
If r = 0 Then
d = d + 1
Picture3.PSet (x, y), vbBlue
Else
Picture3.PSet (x, y), vbRed
End If
c = c + 1
x = x + 15
If x > Picture1.Width Or x > Picture2.Width Then x = 0: y = y + 15
Loop

per = 100 - ((d / c) * 100)

msg = msg & "RESULTS" & vbCrLf
msg = msg & "=============" & vbCrLf
msg = msg & "Total Pixles Scanned: " & c & vbCrLf
msg = msg & "Total Different Pixles: " & d & vbCrLf
msg = msg & "Total Similar Pixles: " & (c - d) & vbCrLf
msg = msg & "Percent Similar: " & per & " %" & vbCrLf
If per > 75 Then
msg = msg & "The computer has judged this as: SIMILAR"
Else
msg = msg & "The computer has judged this as: DIFFERENT"
End If

Form1.Caption = "Image Mapper - Mapped"

MsgBox msg, 0 & 32, "Scan Complete:"
End Sub

Private Sub Command2_Click()
Dim x As Long, y As Long
Dim p1 As Long, p2 As Long
Dim R1 As Long, B1 As Long, G1 As Long
Dim R2 As Long, B2 As Long, G2 As Long
Dim t As Long, r As Integer
Picture3.Cls

t = Text1.Text

Form1.Caption = "Image Mapper - Mapping"

Do Until y > Picture1.Height Or y > Picture2.Height

p1 = Picture1.Point(x, y)
p2 = Picture2.Point(x, y)

RGB_get p1, R1, B1, G1
RGB_get p2, R2, B2, G2

If G1 = -1 Then GoSub 20

Diff r, R1, R2, B1, B2, G1, G2, t
If r = 0 Then
Picture3.PSet (x, y), RGB(R1, B1, G1)
Else
Picture3.PSet (x, y), vbRed
End If

20

x = x + 15
If x > Picture1.Width Or x > Picture2.Width Then x = 0: y = y + 15
Loop

Form1.Caption = "Image Mapper - Mapped"

End Sub

Private Sub Command3_Click()
Dim x As Long, y As Long
Dim p1 As Long, p2 As Long
Dim R1 As Long, B1 As Long, G1 As Long
Dim R2 As Long, B2 As Long, G2 As Long
Dim Ra As Long, Ba As Long, Ga As Long
Picture3.Cls

t = Text1.Text

Form1.Caption = "Image Mapper - Mapping"

Do Until y > Picture1.Height Or y > Picture2.Height

p1 = Picture1.Point(x, y)
p2 = Picture2.Point(x, y)

RGB_get p1, R1, B1, G1
RGB_get p2, R2, B2, G2

Ra = (R1 + R2) / 2
Ba = (B1 + B2) / 2
Ga = (G1 + G2) / 2

If Ga = -1 Then GoSub 10

Picture3.PSet (x, y), RGB(Ra, Ba, Ga)

10

x = x + 15
If x > Picture1.Width Or x > Picture2.Width Then x = 0: y = y + 15
Loop

Form1.Caption = "Image Mapper - Mapped"

End Sub

Private Sub Command4_Click()
On Error GoTo Killer
CommonDialog1.ShowOpen
Picture1 = LoadPicture(CommonDialog1.FileName)
Exit Sub
Killer:
MsgBox "Error!", 0 & 16, "Error:"
End Sub

Private Sub Command5_Click()
On Error GoTo Killer
CommonDialog1.ShowOpen
Picture2 = LoadPicture(CommonDialog1.FileName)
Exit Sub
Killer:
MsgBox "Error!", 0 & 16, "Error:"
End Sub

Private Sub Picture1_Resize()
If Picture1.Width > 3255 Or Picture1.Height > 3255 Then
MsgBox "That picture is too big to be analysed!", 0 & 16, "Error:"
Picture1.Width = 3255
Picture1.Height = 3255
Exit Sub
End If
Picture3.Width = Picture1.Width
Picture3.Height = Picture1.Height
End Sub

Private Sub Picture2_Resize()
If Picture2.Width > 3255 Or Picture12Height > 3255 Then
MsgBox "That picture is too big to be analysed!", 0 & 16, "Error:"
Picture2.Width = 3255
Picture2.Height = 3255
Exit Sub
End If
End Sub

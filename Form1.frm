VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4464
   ClientLeft      =   132
   ClientTop       =   708
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   ScaleHeight     =   4464
   ScaleWidth      =   3900
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3612
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Rich text box used because it holds more characters"
      Top             =   0
      Width           =   3132
      _ExtentX        =   5525
      _ExtentY        =   6371
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2040
      Top             =   840
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2292
      Left            =   3120
      ScaleHeight     =   187
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   2
      Top             =   0
      Width           =   732
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt"
      Height          =   372
      Left            =   1560
      TabIndex        =   1
      Top             =   3600
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "An encryption program that encrypts text to a picture."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   468
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   2616
      WordWrap        =   -1  'True
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Save 
         Caption         =   "Save"
      End
      Begin VB.Menu Load 
         Caption         =   "Load"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This was created to have three leters
'stored in one pixel, so that it does
'not take such a big picture picture to
'hold it, like my last text to picture
'project.It was so bad i didn't upload
'it. Hope this helped you.
'PLEASE VOTE
''''''''''''''''''''''''''''''''''''''
'For better encryption use math to   '
'determine the x and y when plotting '
'and getting points.This was created '
'for learning i did not show examples'
'of better encryption because i want '
'to leave room for ideas so that you '
'can better the encryption and not   '
'have everyone be able to break it   '
'If you have Questions/Comments      '
'E-mail Me at visualcode@juno.com    '
''''''''''''''''''''''''''''''''''''''
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private R 'red
Private G 'green
Private B 'blue

Private char(10000000) As Long
Private text As String

Public Sub get_rgb(color As Long)
B = Int(color / 65536)
G = Int((color - B * 65536) / 256)
R = Int((color - B * 65536) - (G * 256))
'How It Works
'every time you add a red to the rgb
'you are adding 1 to a variable
'
'every time you add a Green to the rgb
'you are adding 256 to a variable
'
'every time you add a Blue to the rgb
'you are adding 65536 to a variable
'
'The variable being Color
'to get the tgb you reverse the process
'as done above
End Sub


Private Sub Command1_Click()
'Clear old text
For i = 0 To 10000000
char(i) = 0
Next i
'Make it start at first character
start = 1
'Clear the picure
Picture1.Cls
'Get the asc of each character
For i = 1 To Len(Text1.text)
char(i) = Asc(Mid(Text1.text, i, 1))
Next i
'Start plotting each point
For x = 0 To Picture1.ScaleWidth - 1
For y = 0 To Picture1.ScaleHeight - 1
'tell it to put the first lettter in g the second in b and third in r
SetPixel Picture1.hdc, x, y, RGB(char(start + 2), char(start), char(start + 1))
start = start + 3
Next y
Next x
'Redraw the picture
Picture1.Refresh
End Sub

Private Sub Command2_Click()
For x = 0 To Picture1.ScaleWidth - 1
For y = 0 To Picture1.ScaleHeight - 1
clr = GetPixel(Picture1.hdc, x, y)
'Get red green and blue values
get_rgb (clr)
'MUST BE G then B then R
'Set up the way it is in encryption
'paste first letter
If G = 0 Then GoTo 1
If G = 255 Then GoTo 1
tmp = tmp & Chr(G)
'paste second letter
If B = 0 Then GoTo 1
If B = 255 Then GoTo 1
tmp = tmp & Chr(B)
'paste third letter
If R = 0 Then GoTo 1
If R = 255 Then GoTo 1
tmp = tmp & Chr(R)
1
Next y
Next x
Text1.text = tmp
End Sub
Private Sub Form_Load()
start = 1
End Sub


Private Sub Form_Resize()
On Error Resume Next
'resize picture
Picture1.Width = Form1.Width - (Text1.Width + 200)
Picture1.Height = Form1.Height - 750
End Sub


Private Sub Load_Click()
'load tthe picture
cd.ShowOpen
Picture1.Picture = LoadPicture(cd.FileName)
End Sub

Private Sub RichTextBox1_Change()
Form1.Caption = "Form1 - " & Len(RichTextBox1)
End Sub

Private Sub Save_Click()
'save the picture
cd.ShowSave
SavePicture Picture1.Image, cd.FileName
End Sub

Private Sub Text1_Change()
'tell the length of text
Form1.Caption = "Form1 - " & Len(Text1.text)
End Sub



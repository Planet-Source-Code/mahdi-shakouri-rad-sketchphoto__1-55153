VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Sketch !"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0000
   ScaleHeight     =   730
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdHelp 
      Caption         =   "About"
      Height          =   495
      Left            =   14280
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdSketch2 
      Caption         =   "Auto Sketch 2"
      Height          =   495
      Left            =   10560
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   495
      Left            =   2040
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin VB.HScrollBar hsDif 
      Height          =   495
      Left            =   3960
      Max             =   125
      Min             =   1
      TabIndex        =   8
      Top             =   120
      Value           =   10
      Width           =   1695
   End
   Begin VB.CommandButton cmdSketch 
      Caption         =   "Auto Sketch 1"
      Height          =   495
      Left            =   8280
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "Clear"
      Height          =   495
      Left            =   13320
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   11880
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpenFit 
      Caption         =   "Open and Fit"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   7680
      ScaleHeight     =   122
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      Top             =   720
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   10215
      Left            =   120
      MouseIcon       =   "Form1.frx":0442
      ScaleHeight     =   681
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   493
      TabIndex        =   0
      Top             =   720
      Width           =   7395
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   10
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblDif 
      Caption         =   "Dif=10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblColor2 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   7680
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblColor1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w0, h0, w, h As Single

Private Sub cmdCls_Click()
  Picture2.Cls
  Picture2.Picture = LoadPicture()
End Sub

Private Sub cmdHelp_Click()
  Form3.Show
End Sub

Private Sub cmdOpen_Click()
  CommonDialog1.CancelError = True
  On Error GoTo ja
  
  CommonDialog1.Filter = "Image|*.bmp;*.gif;*.jpg"
  CommonDialog1.ShowOpen
  ' open picture in Picture2 and then fit itin Picture1!
  Picture1.Picture = LoadPicture(CommonDialog1.FileName)
  Picture2.Picture = LoadPicture(CommonDialog1.FileName)
  
  w0 = Picture2.Width
  h0 = Picture2.Height
  
  w = 493
  h = 681
  
  If w0 < w Then w = w0
  If h0 < h Then h = h0
  
  Picture1.Width = w
  Picture1.Height = h
  
  Picture2.Width = Picture1.Width
  Picture2.Height = Picture1.Height
  
  'Picture1.Picture = LoadPicture()
  'Picture1.PaintPicture Picture2.Picture, 0, 0, Picture1.Width, Picture1.Height, 0, 0, w0, h0, vbSrcCopy
  Picture2.Picture = LoadPicture()
  
ja:
End Sub

Private Sub cmdOpenFit_Click()
  CommonDialog1.CancelError = True
  On Error GoTo ja
  
  CommonDialog1.Filter = "Image|*.bmp;*.gif;*.jpg"
  CommonDialog1.ShowOpen
  ' open picture in Picture2 and then fit itin Picture1!
  Picture1.Picture = LoadPicture(CommonDialog1.FileName)
  Picture2.Picture = LoadPicture(CommonDialog1.FileName)
  ' Fit
  w0 = Picture2.Width
  h0 = Picture2.Height
  
  w = 493
  h = 681
  
  If w0 / h0 > w / h Then
    ' resize based on width
    Picture1.Width = w
    Picture1.Height = w * h0 / w0
    Else
    ' resize based on height
    Picture1.Height = h
    Picture1.Width = h * w0 / h0
  End If
  Picture2.Width = Picture1.Width
  Picture2.Height = Picture1.Height
  
  Picture1.Picture = LoadPicture()
  Picture1.PaintPicture Picture2.Picture, 0, 0, Picture1.Width, Picture1.Height, 0, 0, w0, h0, vbSrcCopy
  Picture2.Picture = LoadPicture()
  
ja:
End Sub

Private Sub cmdSave_Click()
  Dim FName As String
  CommonDialog1.CancelError = True
  On Error GoTo ja
  CommonDialog1.Filter = "*.jpg"
  CommonDialog1.ShowSave
  FName = CommonDialog1.FileName
  If Right$(FName, 4) <> ".jpg" Then FName = FName + ".jpg"
  SavePicture Picture2.Image, FName
ja:
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then
    lblColor2.BackColor = Picture1.Point(x, y)
  End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If x < 0 Or x >= Picture1.ScaleWidth Or y < 0 Or y >= Picture1.ScaleHeight Then Exit Sub
  lblColor1.BackColor = Picture1.Point(x, y)
  
  If Button = vbLeftButton Then
    Picture2.ForeColor = lblColor1.BackColor
    Picture2.PSet (x, y)
  End If

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then
    Picture2.ForeColor = lblColor2.BackColor
    Picture2.PSet (x, y)
  End If
End Sub

Private Function BW(c As Long) As Integer
Dim R, G, B As Integer
  R = c Mod 256
  G = (c \ 256) Mod 256
  B = (c \ 256 \ 256) Mod 256
  BW = (R + G + B) / 3
End Function

Private Sub cmdSketch_Click()
Dim x, y As Integer
Dim c1, c2 As Integer
Dim total, done As Long
Picture2.Cls
total = Picture1.Width + Picture1.Height
done = 0
' in Y direction :
For x = 0 To Picture1.Width - 1
  done = done + 1
  lblPercent = Str$(Int(100 * done / total)) + "%"
  DoEvents
  For y = 0 To Picture1.Height - 2
    c1 = (BW(Picture1.Point(x, y)))
    c2 = (BW(Picture1.Point(x, y + 1)))
    If Abs(c1 - c2) > hsDif.Value Then
      Picture2.PSet (x, y), vbBlack
    End If
  Next y
Next x
' in X direction :
For y = 0 To Picture1.Height - 1
  done = done + 1
  lblPercent = Str$(Int(100 * done / total)) + "%"
  DoEvents
  For x = 0 To Picture1.Width - 2
    c1 = (BW(Picture1.Point(x, y)))
    c2 = (BW(Picture1.Point(x + 1, y)))
    If Abs(c1 - c2) > hsDif.Value Then
      Picture2.PSet (x, y), vbBlack
    End If
  Next x
Next y
lblPercent = "%"
End Sub
Private Sub cmdSketch2_Click()
Dim x, y As Integer
Dim c1, c2 As Integer
Dim c As Long
Dim total, done As Long
Picture2.Cls
total = Picture1.Width
done = 0
' in Y direction :
c = Picture1.Point(0, 0)
For x = 0 To Picture1.Width - 1
  done = done + 1
  lblPercent = Str$(Int(100 * done / total)) + "%"
  DoEvents
  For y = 0 To Picture1.Height - 2
    If BW(Picture1.Point(x, y)) > hsDif.Value Then
      c = vbWhite
      Else
      c = vbBlack
    End If
    Picture2.PSet (x, y), c
  Next y
Next x
' in X direction :
'For y = 0 To Picture1.Height - 1
'  c = Picture1.Point(0, y)
'  done = done + 1
'  lblPercent = Str$(Int(100 * done / total)) + "%"
'  DoEvents
'  For x = 0 To Picture1.Width - 2
'    c1 = (BW(Picture1.Point(x, y)))
'    c2 = (BW(Picture1.Point(x + 1, y)))
'    If Abs(c1 - c2) > hsDif.Value Then
'      c = Picture1.Point(x, y)
'    End If
'    Picture2.PSet (x, y), c
'  Next x
'Next y
lblPercent = "%"
End Sub

Private Sub hsDif_Change()
  lblDif = "Dif=" & Str$(hsDif.Value)
End Sub


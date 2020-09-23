VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text Picture make"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15420
   Icon            =   "Text String.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   15420
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   14631
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "wmf to text"
      TabPicture(0)   =   "Text String.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Image3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Image4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Image5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Image6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Timer1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Pil"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "P1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command5"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command6"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command7"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Command8"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "R1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Check1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Command10"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "List1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Command17"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "String to text picture"
      TabPicture(1)   =   "Text String.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Command11"
      Tab(1).Control(3)=   "P2"
      Tab(1).Control(4)=   "Text1"
      Tab(1).Control(5)=   "Command2"
      Tab(1).Control(6)=   "R2"
      Tab(1).Control(7)=   "Command9"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Make your own picture or load a picture"
      TabPicture(2)   =   "Text String.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape1"
      Tab(2).Control(1)=   "Shape2"
      Tab(2).Control(2)=   "Shape5"
      Tab(2).Control(3)=   "Shape6"
      Tab(2).Control(4)=   "Label4"
      Tab(2).Control(5)=   "P4"
      Tab(2).Control(6)=   "Command13"
      Tab(2).Control(7)=   "Command12"
      Tab(2).Control(8)=   "Frame1"
      Tab(2).Control(9)=   "HScroll1"
      Tab(2).Control(10)=   "P3"
      Tab(2).Control(11)=   "Command14"
      Tab(2).Control(12)=   "Command15"
      Tab(2).Control(13)=   "loadpic"
      Tab(2).Control(14)=   "Command16"
      Tab(2).ControlCount=   15
      Begin VB.CommandButton Command17 
         Caption         =   "Command17"
         Height          =   510
         Left            =   5805
         TabIndex        =   36
         Top             =   7740
         Width           =   555
      End
      Begin VB.ListBox List1 
         Height          =   450
         ItemData        =   "Text String.frx":091E
         Left            =   6435
         List            =   "Text String.frx":0925
         Sorted          =   -1  'True
         TabIndex        =   35
         Top             =   7785
         Width           =   2130
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Save to ""vmf to text Tab"""
         Height          =   285
         Left            =   -63480
         TabIndex        =   34
         Top             =   7515
         Width           =   2040
      End
      Begin MSComDlg.CommonDialog loadpic 
         Left            =   -66720
         Top             =   7605
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Open wmf"
         FileName        =   "*.wmf"
         Filter          =   ".wmf"
         InitDir         =   "app.path & ""\"""
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Load picture"
         Height          =   285
         Left            =   -63480
         TabIndex        =   33
         Top             =   7200
         Width           =   2040
      End
      Begin VB.CommandButton Command14 
         Caption         =   "clear"
         Height          =   285
         Left            =   -63480
         TabIndex        =   32
         Top             =   7830
         Width           =   1500
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Save"
         Height          =   330
         Left            =   3375
         TabIndex        =   27
         Top             =   7920
         Width           =   825
      End
      Begin VB.CheckBox Check1 
         Caption         =   "DoEvents"
         Height          =   285
         Left            =   2205
         TabIndex        =   26
         Top             =   7920
         Width           =   1050
      End
      Begin VB.TextBox R1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   12
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4290
         Left            =   4365
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   25
         Top             =   495
         Width           =   4290
      End
      Begin VB.CommandButton Command8 
         Caption         =   "6"
         Height          =   285
         Left            =   1845
         TabIndex        =   24
         Top             =   7920
         Width           =   240
      End
      Begin VB.CommandButton Command7 
         Caption         =   "5"
         Height          =   285
         Left            =   1485
         TabIndex        =   23
         Top             =   7920
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "4"
         Height          =   285
         Left            =   1125
         TabIndex        =   22
         Top             =   7920
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "3"
         Height          =   285
         Left            =   765
         TabIndex        =   21
         Top             =   7920
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "2"
         Height          =   285
         Left            =   405
         TabIndex        =   20
         Top             =   7920
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "1"
         Height          =   285
         Left            =   90
         TabIndex        =   19
         Top             =   7920
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Pic to text"
         Height          =   330
         Left            =   4275
         TabIndex        =   18
         Top             =   7920
         Width           =   915
      End
      Begin VB.PictureBox P1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   4260
         Left            =   90
         Picture         =   "Text String.frx":092D
         ScaleHeight     =   17.5
         ScaleMode       =   4  'Character
         ScaleWidth      =   35
         TabIndex        =   17
         Top             =   495
         Width           =   4260
      End
      Begin VB.CommandButton Command9 
         Caption         =   "save"
         Height          =   375
         Left            =   -68070
         TabIndex        =   16
         Top             =   6885
         Width           =   1050
      End
      Begin VB.TextBox R2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   12
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2535
         Left            =   -74910
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   15
         Top             =   3960
         Width           =   15180
      End
      Begin VB.CommandButton Command2 
         Caption         =   "String to text"
         Height          =   375
         Left            =   -68070
         TabIndex        =   14
         Top             =   7290
         Width           =   1050
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74910
         MaxLength       =   10
         TabIndex        =   13
         Top             =   6885
         Width           =   6720
      End
      Begin VB.PictureBox P2 
         Height          =   2895
         Left            =   -74910
         ScaleHeight     =   5.001
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   26.67
         TabIndex        =   12
         Top             =   495
         Width           =   15180
      End
      Begin VB.PictureBox P3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawStyle       =   3  'Dash-Dot
         DrawWidth       =   10
         Height          =   6270
         Left            =   -74775
         ScaleHeight     =   414
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   444
         TabIndex        =   11
         Top             =   540
         Width           =   6720
         Begin VB.Shape Shape3 
            DrawMode        =   6  'Mask Pen Not
            Height          =   375
            Left            =   720
            Shape           =   2  'Oval
            Top             =   810
            Width           =   375
            Visible         =   0   'False
         End
         Begin VB.Shape Shape4 
            DrawMode        =   6  'Mask Pen Not
            Height          =   255
            Left            =   90
            Top             =   1440
            Width           =   375
            Visible         =   0   'False
         End
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   285
         Left            =   -74640
         Max             =   20
         Min             =   10
         TabIndex        =   10
         Top             =   7290
         Value           =   10
         Width           =   1365
      End
      Begin VB.Frame Frame1 
         Caption         =   "Draw tools"
         Height          =   1095
         Left            =   -72795
         TabIndex        =   6
         Top             =   7065
         Width           =   1365
         Begin VB.OptionButton Option1 
            Caption         =   "Pen"
            Height          =   195
            Left            =   135
            TabIndex        =   9
            Top             =   225
            Width           =   1005
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Circle"
            Height          =   195
            Left            =   135
            TabIndex        =   8
            Top             =   495
            Width           =   915
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Square"
            Height          =   240
            Left            =   135
            TabIndex        =   7
            Top             =   765
            Width           =   960
         End
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Clear"
         Height          =   330
         Left            =   -68070
         TabIndex        =   5
         Top             =   7695
         Width           =   1050
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Save to vmf to text Tab"
         Height          =   285
         Left            =   -71175
         TabIndex        =   4
         Top             =   7200
         Width           =   2040
      End
      Begin VB.CommandButton Command13 
         Caption         =   "clear"
         Height          =   285
         Left            =   -71175
         TabIndex        =   3
         Top             =   7650
         Width           =   1500
      End
      Begin VB.PictureBox Pil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   5130
         ScaleHeight     =   555
         ScaleWidth      =   780
         TabIndex        =   2
         Top             =   7335
         Width           =   780
         Visible         =   0   'False
         Begin VB.Line Line1 
            BorderWidth     =   5
            X1              =   495
            X2              =   90
            Y1              =   45
            Y2              =   450
         End
         Begin VB.Line Line2 
            BorderWidth     =   5
            X1              =   270
            X2              =   45
            Y1              =   495
            Y2              =   495
         End
         Begin VB.Line Line3 
            BorderWidth     =   5
            X1              =   45
            X2              =   45
            Y1              =   270
            Y2              =   450
         End
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   5985
         Top             =   5445
      End
      Begin VB.PictureBox P4 
         Height          =   6315
         Left            =   -67755
         ScaleHeight     =   26.063
         ScaleMode       =   4  'Character
         ScaleWidth      =   59.25
         TabIndex        =   1
         Top             =   495
         Width           =   7170
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   285
         Left            =   7245
         TabIndex        =   37
         Top             =   7470
         Width           =   1365
      End
      Begin VB.Image Image6 
         Height          =   4200
         Left            =   990
         Picture         =   "Text String.frx":3ECB
         Top             =   3690
         Width           =   4200
         Visible         =   0   'False
      End
      Begin VB.Image Image5 
         Height          =   4890
         Left            =   8280
         Picture         =   "Text String.frx":6701
         Top             =   2385
         Width           =   3510
         Visible         =   0   'False
      End
      Begin VB.Image Image4 
         Height          =   4875
         Left            =   1440
         Picture         =   "Text String.frx":6DA3
         Top             =   1620
         Width           =   3045
         Visible         =   0   'False
      End
      Begin VB.Image Image3 
         Height          =   3915
         Left            =   270
         Picture         =   "Text String.frx":71E5
         Top             =   900
         Width           =   7995
         Visible         =   0   'False
      End
      Begin VB.Image Image2 
         Height          =   4755
         Left            =   7425
         Picture         =   "Text String.frx":76E7
         Top             =   2745
         Width           =   2790
         Visible         =   0   'False
      End
      Begin VB.Image Image1 
         Height          =   5040
         Left            =   180
         Picture         =   "Text String.frx":7B09
         Top             =   1935
         Width           =   4515
         Visible         =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   240
         Left            =   315
         TabIndex        =   31
         Top             =   3915
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "Text Picture"
         Height          =   240
         Left            =   -74910
         TabIndex        =   30
         Top             =   3690
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Write a String (max 10 characters)"
         Height          =   240
         Left            =   -74910
         TabIndex        =   29
         Top             =   6615
         Width           =   2940
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "DrawWidth: 10"
         Height          =   195
         Left            =   -74640
         TabIndex        =   28
         Top             =   7065
         Width           =   1410
      End
      Begin VB.Shape Shape6 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   -71625
         Top             =   1935
         Width           =   375
      End
      Begin VB.Shape Shape5 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   -70905
         Top             =   2070
         Width           =   375
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   -73065
         Top             =   1845
         Width           =   375
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   -72345
         Top             =   1665
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim j As Integer
Dim c As Integer
Dim l As Integer
Dim kludda As Boolean
Dim lod As Integer
Dim vag As Integer
Dim XX As Double, YY As Double
Dim XX2 As Double, YY2 As Double
Dim isThere As Boolean
Dim w As Integer
Dim e As Integer
Dim nColor As Long
Dim nVector(8001) As Long
Dim a As Long
Dim Blink As Byte
Dim s As Byte
Private Sub Command1_Click()
Timer1.Enabled = False
Pil.Visible = False
P1.BackColor = vbWhite

For i = 0 To P1.ScaleHeight
    For j = 1 To P1.ScaleWidth - 1
        If P1.Point(j, i) = vbWhite Then
            R1.Text = R1.Text & "+"
        ElseIf P1.Point(j, i) = vbBlack Then
            R1.Text = R1.Text & "#"
        End If
        If Check1.Value = Checked Then
        DoEvents
        End If
    Next j
    R1.Text = R1.Text & vbCrLf
    
Next i
End Sub

Private Sub Command10_Click()
    Open App.Path & "\picture.txt" For Append As #1
    Print #1, R1
    Close #1
End Sub

Private Sub Command11_Click()
P2.ScaleMode = 7
R2 = ""
P2.Cls
Text1 = ""
c = 0
s = 0
End Sub

Private Sub Command12_Click()
P1.AutoSize = True
Clipboard.Clear
Clipboard.SetData P3.Image
P1.Picture = Clipboard.GetData(wmf)
R1 = ""
R1.Left = P1.Left + P1.Width
R1.Height = P1.Height
Timer1.Enabled = True
SSTab1.Tab = 0
End Sub

Private Sub Command13_Click()
P3.Cls
End Sub

Private Sub Command14_Click()
P4.Cls
End Sub

Private Sub Command15_Click()
On Error Resume Next
loadpic.ShowOpen
P4.Picture = LoadPicture(loadpic.FileName)
End Sub

Private Sub Command16_Click()
P1.AutoSize = False
P1.Width = P4.Width
P1.Height = P4.Height
P1.Picture = P4.Picture
R1 = ""
R1.Left = P1.Left + P1.Width
R1.Height = P1.Height
Timer1.Enabled = True
SSTab1.Tab = 0
End Sub

Private Sub Command17_Click()

e = 0
For i = 0 To P1.ScaleHeight
    For j = 0 To P1.ScaleWidth
      nColor = P1.Point(j, i)
      isThere = False
      For w = 0 To e
        If nColor = nVector(w) Then
             isThere = True
        End If
      Next w
      e = e + 1
    If isThere = False Then
        nVector(e) = nColor
        List1.AddItem (nColor)
    End If
    Next j
    
Next i
'For i = 0 To e
'List1.AddItem (nVector(e))
'Next i

End Sub

Private Sub Command18_Click()

End Sub

Private Sub Command2_Click()
P2.ScaleMode = 4
R2 = ""

For i = 1 To P2.ScaleHeight
    For j = 1 To c * 5
        If P2.Point(j, i) <> vbWhite Then
            R2.Text = R2.Text & "+"
        ElseIf P2.Point(j, i) = vbWhite Then
            R2.Text = R2.Text & "#"
        End If
    Next j
    R2.Text = R2.Text & vbCrLf
Next i


End Sub



Private Sub Command3_Click()
P1.AutoSize = True
P1.Picture = Image1
R1 = ""
R1.Left = P1.Left + P1.Width
R1.Height = P1.Height



End Sub

Private Sub Command4_Click()
P1.AutoSize = True
P1.Picture = Image2
R1 = ""
R1.Left = P1.Left + P1.Width
R1.Height = P1.Height


End Sub

Private Sub Command5_Click()
P1.AutoSize = True
P1.Picture = Image3
R1 = ""
R1.Left = P1.Left + P1.Width
R1.Height = P1.Height


End Sub

Private Sub Command6_Click()
P1.AutoSize = True
P1.Picture = Image4
R1 = ""
R1.Left = P1.Left + P1.Width
R1.Height = P1.Height


End Sub

Private Sub Command7_Click()
P1.AutoSize = True
P1.Picture = Image5
R1 = ""
R1.Left = P1.Left + P1.Width
R1.Height = P1.Height

End Sub

Private Sub Command8_Click()
P1.AutoSize = True
P1.Picture = Image6
R1 = ""
R1.Left = P1.Left + P1.Width
R1.Height = P1.Height


End Sub

Private Sub Command9_Click()
    Open App.Path & "\stringpic.txt" For Append As #1
    Print #1, R2
    Close #1
End Sub

Private Sub Form_Load()
c = 0
s = 0
P2.BackColor = vbWhite
loadpic.InitDir = App.Path & "\"
End Sub

Private Sub HScroll1_Change()
Label4 = "DrawWidth: " & HScroll1.Value
P3.DrawWidth = HScroll1.Value
End Sub

Private Sub List1_Click()
Label5 = List1
End Sub

Private Sub P1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1 = P1.Point(X, Y)
End Sub

Private Sub P3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
kludda = True
vag = X
lod = Y
'***Square***'
If kludda = True And Option3.Value = True Then
     XX = X: YY = Y
     XX2 = X: YY2 = Y
     Shape4.Visible = True
     Shape4.Left = X:
     Shape4.Width = 0: Shape4.Height = 0

End If

'***oval***'
 If kludda = True And Option2.Value = True Then
    XX = X: YY = Y
    XX2 = X: YY2 = Y
    Shape3.Visible = True
    Shape3.Left = X:
    Shape3.Width = 0: Shape3.Height = 0

End If
End Sub

Private Sub P3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If kludda And Option1 = True Then
    P3.Line (X, Y)-(vag, lod), vbBlack
     vag = X
    lod = Y
End If
'***Rektangle***'
If kludda = True And Option3 = True Then
        XX2 = X: YY2 = Y
        Shape4.Left = IIf(X > XX, XX, X)
        Shape4.Top = IIf(Y > YY, YY, Y)
        Shape4.Width = Abs(X - XX)
        Shape4.Height = Abs(Y - YY)
End If

'***cirkel***'
If kludda = True And Option2 = True Then
        XX2 = X: YY2 = Y
        Shape3.Left = IIf(X > XX, XX, X)
        Shape3.Top = IIf(Y > YY, YY, Y)
        Shape3.Width = Abs(X - XX)
        Shape3.Height = Abs(Y - YY)
End If
End Sub

Private Sub P3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 kludda = False
 '***Rektangle***'
On Error Resume Next

If Option3 = True Then
    P3.DrawWidth = HScroll1.Value
    If XX2 <> XX Then P3.Line ((XX), (YY))-(XX2, YY2), vbBlack, B
        Shape4.Visible = False
        P3.Line (XX, YY)-(XX2, YY2), vbBlack, B
    End If
    
    
'***oval***'
If Option2 = True Then
    P3.DrawWidth = HScroll1.Value
    rad = IIf(Abs(YY2 - YY) > Abs(XX2 - XX), Abs(YY2 - YY) / 2, Abs(XX2 - XX) / 2)
    If XX2 <> XX Then P3.Circle ((XX2 + XX) / 2, (YY2 + YY) / 2), rad, Farg, , , Abs(YY2 - YY) / Abs(XX2 - XX)
    Shape3.Visible = False
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
s = s + 1
If s <= 10 Then
Select Case KeyCode
    Case vbKeyA
        P2.PaintPicture LoadPicture(App.Path & "\a.wmf"), c, 0, 3, 5
        c = c + 3
        
    Case vbKeyB
        P2.PaintPicture LoadPicture(App.Path & "\b.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyC
        P2.PaintPicture LoadPicture(App.Path & "\c.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyD
        P2.PaintPicture LoadPicture(App.Path & "\d.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyE
        P2.PaintPicture LoadPicture(App.Path & "\e.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyF
        P2.PaintPicture LoadPicture(App.Path & "\f.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyG
        P2.PaintPicture LoadPicture(App.Path & "\g.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyH
        P2.PaintPicture LoadPicture(App.Path & "\h.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyI
        P2.PaintPicture LoadPicture(App.Path & "\i.wmf"), c, 0, 3, 5
        c = c + 1
    Case vbKeyJ
        P2.PaintPicture LoadPicture(App.Path & "\j.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyK
        P2.PaintPicture LoadPicture(App.Path & "\k.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyL
        P2.PaintPicture LoadPicture(App.Path & "\l.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyM
        P2.PaintPicture LoadPicture(App.Path & "\m.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyN
        P2.PaintPicture LoadPicture(App.Path & "\n.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyO
        P2.PaintPicture LoadPicture(App.Path & "\o.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyP
        P2.PaintPicture LoadPicture(App.Path & "\p.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyQ
        P2.PaintPicture LoadPicture(App.Path & "\q.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyR
        P2.PaintPicture LoadPicture(App.Path & "\r.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyS
        P2.PaintPicture LoadPicture(App.Path & "\s.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyT
        P2.PaintPicture LoadPicture(App.Path & "\t.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyY
        P2.PaintPicture LoadPicture(App.Path & "\y.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyU
        P2.PaintPicture LoadPicture(App.Path & "\u.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyW
        P2.PaintPicture LoadPicture(App.Path & "\w.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyV
        P2.PaintPicture LoadPicture(App.Path & "\v.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyY
        P2.PaintPicture LoadPicture(App.Path & "\y.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyX
        P2.PaintPicture LoadPicture(App.Path & "\x.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeyZ
        P2.PaintPicture LoadPicture(App.Path & "\z.wmf"), c, 0, 3, 5
        c = c + 3
    Case vbKeySpace
        P2.PaintPicture LoadPicture(App.Path & "\32.wmf"), c, 0, 3, 5
        c = c + 1
    Case Else
        Beep
    End Select
End If
End Sub
Private Sub Timer1_Timer()
Blink = Blink + 1
If Blink = 2 Then
    Pil.Visible = False
Else
    Pil.Visible = True
    Blink = 1
End If

End Sub
Sub count_color()

End Sub

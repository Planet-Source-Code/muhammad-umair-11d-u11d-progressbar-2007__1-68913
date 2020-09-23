VERSION 5.00
Begin VB.Form FrmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TT 
      Interval        =   100
      Left            =   3000
      Top             =   1860
   End
   Begin VB.Timer T 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   1860
   End
   Begin U11DProgressBar.ProgressBar U 
      Height          =   270
      Left            =   135
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   476
      Max             =   250
      Value           =   0
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "U"
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Const LWA_ALPHA = &H2
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Sub Form_Load()
TOPFORM Me.hWnd, True
DrawBackground
FormFadeIn Me, 0, 240, 4
'LOGO
End Sub
Private Sub RoundRectBorder(nObject As Object, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, X3 As Long, Y3 As Long, nColor As ColorConstants)
Dim A As Variant
A = nObject.ForeColor
nObject.ForeColor = nColor
RoundRect nObject.hDC, X1, Y1, X2, Y2, X3, Y3
nObject.ForeColor = A
End Sub
Private Sub TOPFORM(hWnd As Long, Action As Boolean)
If Action = True Then
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
Else
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End If
End Sub
Private Sub FormFadeIn(ByRef nForm As Form, Optional ByVal nFadeStart As Byte = 0, Optional ByVal nFadeEnd As Byte = 255, Optional ByVal nFadeInSpeed As Byte = 5)
Dim c
Dim ne As Integer, en(32767) As Boolean
For Each c In nForm.Controls
 ne = ne + 1
 en(ne) = c.Enabled
 c.Enabled = False
Next
If nFadeEnd = 0 Then
    nFadeEnd = 255
End If
If nFadeInSpeed = 0 Then
    nFadeInSpeed = 5
End If
If nFadeStart >= nFadeEnd Then
    nFadeStart = 0
ElseIf nFadeEnd <= nFadeStart Then
    nFadeEnd = 255
End If

   TransparentsForm nForm.hWnd, 0
    nForm.Show
    Dim I As Long
    For I = nFadeStart To nFadeEnd Step nFadeInSpeed
        TransparentsForm nForm.hWnd, CByte(I)
        DoEvents
        Call Sleep(5)
    Next
    TransparentsForm nForm.hWnd, nFadeEnd
    I = 0
For Each c In nForm.Controls
 I = I + 1
 c.Enabled = en(I)
Next
End Sub
Private Function TransparentsForm(FormhWnd As Long, Alpha As Byte) As Boolean
    SetWindowLong FormhWnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes FormhWnd, 0, Alpha, LWA_ALPHA
End Function
Private Sub FormFadeOut(ByRef nForm As Form)
On Local Error Resume Next
Dim c
Dim S As Integer
For Each c In nForm.Controls
 c.Enabled = False
Next

Dim I As Long
    For I = 240 To 0 Step -5
        TransparentsForm nForm.hWnd, CByte(I)
        DoEvents
        Call Sleep(5)
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
FormFadeOut Me
TOPFORM Me.hWnd, False
End Sub
Private Sub DrawBackground()
Dim Lonrect As Long
Lonrect = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 20, 20)
SetWindowRgn Me.hWnd, Lonrect, True
RoundRectBorder Me, 0, 0, Me.ScaleWidth - 1, Me.ScaleHeight - 1, 20, 20, vbWhite

        Me.BackColor = vbWhite
        U.DrawGradientFourColour Me.hDC, 3, 3, Me.ScaleWidth - 7, Me.ScaleHeight / 2 - 1, RGB(231, 243, 224), RGB(220, 234, 244), RGB(122, 183, 222), RGB(122, 183, 222)
        U.DrawGradientFourColour Me.hDC, 3, Me.ScaleHeight / 2 - 1, Me.ScaleWidth - 7, Me.ScaleHeight / 2 - 3, RGB(78, 153, 200), RGB(78, 153, 200), RGB(220, 234, 244), RGB(239, 209, 234)
        Me.ForeColor = RGB(161, 189, 207)
        RoundRect Me.hDC, 0, 0, ScaleWidth - 1, ScaleHeight - 1, 20, 20
        Me.ForeColor = RGB(255, 255, 255)
        RoundRect Me.hDC, 1, 1, ScaleWidth - 2, ScaleHeight - 2, 20, 20
        RoundRect Me.hDC, 2, 2, ScaleWidth - 3, ScaleHeight - 3, 20, 20
        
        Me.PSet (3, 4), RGB(255, 255, 255)
        Me.PSet (4, 3), RGB(255, 255, 255)
        Me.PSet (3, 6), RGB(255, 255, 255)
        Me.PSet (4, 5), RGB(255, 255, 255)
        Me.PSet (5, 4), RGB(255, 255, 255)
        Me.PSet (7, 3), RGB(255, 255, 255)
        
        Me.PSet (3, Me.ScaleHeight - 6), RGB(255, 255, 255)
        Me.PSet (4, Me.ScaleHeight - 5), RGB(255, 255, 255)
        Me.PSet (3, Me.ScaleHeight - 8), RGB(255, 255, 255)
        Me.PSet (4, Me.ScaleHeight - 7), RGB(255, 255, 255)
        Me.PSet (5, Me.ScaleHeight - 6), RGB(255, 255, 255)
        Me.PSet (7, Me.ScaleHeight - 5), RGB(255, 255, 255)
        
        Me.PSet (Me.ScaleWidth - 5, 4), RGB(255, 255, 255)
        Me.PSet (Me.ScaleWidth - 6, 3), RGB(255, 255, 255)
        Me.PSet (Me.ScaleWidth - 5, 6), RGB(255, 255, 255)
        Me.PSet (Me.ScaleWidth - 6, 5), RGB(255, 255, 255)
        Me.PSet (Me.ScaleWidth - 7, 4), RGB(255, 255, 255)
        Me.PSet (Me.ScaleWidth - 9, 3), RGB(255, 255, 255)
    
        Me.PSet (Me.ScaleWidth - 5, Me.ScaleHeight - 6), RGB(255, 255, 255)
        Me.PSet (Me.ScaleWidth - 6, Me.ScaleHeight - 5), RGB(255, 255, 255)
        Me.PSet (Me.ScaleWidth - 5, Me.ScaleHeight - 8), RGB(255, 255, 255)
        Me.PSet (Me.ScaleWidth - 6, Me.ScaleHeight - 7), RGB(255, 255, 255)
        Me.PSet (Me.ScaleWidth - 7, Me.ScaleHeight - 6), RGB(255, 255, 255)
        Me.PSet (Me.ScaleWidth - 9, Me.ScaleHeight - 5), RGB(255, 255, 255)
        
        Me.Refresh
End Sub

        


Private Sub LOGO()
Dim PREFontName As String, PREFontSize As Integer, PREFontCOLOR As Long
Dim PREFontBold As Boolean, PREFontItalic As Boolean, PREFontStrikethru As Boolean, PREFontUnderline As Boolean
Dim I As Long

TT.Enabled = False
PREFontName = Me.FontName
PREFontSize = Me.FontSize
PREFontCOLOR = Me.ForeColor
PREFontBold = Me.FontBold
PREFontItalic = Me.FontItalic
PREFontStrikethru = Me.FontStrikethru
PREFontUnderline = Me.FontUnderline



Me.FontName = "Tahoma"
Me.FontSize = 55
Me.ForeColor = RGB(255, 255, 255)
Me.FontBold = True
Me.CurrentX = 14
Me.CurrentY = 60
TEXTShadow "UMAIR_11D", RGB(229, 237, 247)
Me.Refresh

Sleep 500

Me.CurrentX = Me.ScaleWidth - 25
Me.CurrentY = 70
Me.FontSize = 12
Me.Print "®"
Me.Refresh

Sleep 500

Me.CurrentX = 120
Me.CurrentY = 150
Me.FontSize = 30
Me.ForeColor = RGB(255, 255, 255)


For I = 0 To 8
    Me.CurrentX = 120
    Me.CurrentY = 150
    TEXTShadow Mid("PRESENTS", 1, CByte(I)), RGB(190, 210, 234)
    DoEvents
    Call Sleep(20)
Next

Sleep 2000

Me.Cls
DrawBackground


Me.FontSize = 38


For I = 0 To 16
    Me.CurrentX = 5
    Me.CurrentY = 65
    TEXTShadow Mid("U11D ProgressBar", 1, CByte(I)), RGB(229, 237, 247)
    DoEvents
    Call Sleep(20)
Next


Sleep 500

Me.CurrentX = Me.ScaleWidth - 25
Me.CurrentY = 65
Me.FontSize = 15
Me.Print "™"
Me.Refresh

Sleep 100

Me.CurrentX = 10
Me.CurrentY = 125
Me.FontSize = 10
Me.Print "RELIABLE , FLEXIBLE , COMPATIBLE , FASTER , POWERFUL , EASIER TO USE"
Me.Refresh


Sleep 1000

Me.ForeColor = vbWhite
Me.FontSize = 10

For I = 0 To 65
    Me.CurrentX = 12
    Me.CurrentY = 145
    Me.Print Mid("U11D ProgressBar is Very Quick, Powerful & New styles ProgressBar.", 1, CByte(I))
    DoEvents
    Call Sleep(10)
Next
For I = 0 To 67
    Me.CurrentX = 12
    Me.CurrentY = 160
    Me.Print Mid("U11D ProgressBar enables you to customize the appearance of your", 1, CByte(I))
    DoEvents
    Call Sleep(10)
Next
For I = 0 To 44
    Me.CurrentX = 12
    Me.CurrentY = 175
    Me.Print Mid("applications to suit your individual needs.", 1, CByte(I))
    DoEvents
    Call Sleep(10)
Next

Me.FontSize = 8

For I = 0 To 30
    Me.CurrentX = 12
    Me.CurrentY = 200
    Me.Print Mid("If You Find Any Problems/Bug.", 1, CByte(I))
    DoEvents
    Call Sleep(10)
Next
Call Sleep(100)

For I = 0 To 32
    Me.CurrentX = 12
    Me.CurrentY = 210
    Me.Print Mid("Any Questions For This Project.", 1, CByte(I))
    DoEvents
    Call Sleep(10)
Next

Call Sleep(100)

For I = 0 To 26
    Me.CurrentX = 12
    Me.CurrentY = 225
    Me.Print Mid("Email:Umair_11D@Yahoo.com", 1, CByte(I))
    DoEvents
    Call Sleep(10)
Next
Call Sleep(100)

For I = 0 To 41
    Me.CurrentX = 12
    Me.CurrentY = 237
    Me.Print Mid("Voice NO :+923453021375 , +923222678852", 1, CByte(I))
    DoEvents
    Call Sleep(10)
Next

Sleep 500

Me.FontName = PREFontName
Me.FontSize = PREFontSize
Me.ForeColor = PREFontCOLOR
Me.FontBold = PREFontBold
Me.FontItalic = PREFontItalic
Me.FontStrikethru = PREFontStrikethru
Me.FontUnderline = PREFontUnderline

U.Visible = True
T.Enabled = True
End Sub
Private Sub TEXTShadow(STR As String, ShadowColor As Long)
Dim PREColor As Long
Dim OW As Long, OH As Long
PREColor = Me.ForeColor
OW = Me.CurrentX
OH = Me.CurrentY

        Me.ForeColor = ShadowColor
        Me.CurrentX = OW + 1
        Me.CurrentY = OH
        Me.Print STR
        Me.CurrentX = OW - 1
        Me.CurrentY = OH
        Me.Print STR
        Me.CurrentY = OH - 1
        Me.CurrentX = OW
        Me.Print STR
        Me.CurrentY = OH + 1
        Me.CurrentX = OW
        Me.Print STR
        Me.ForeColor = PREColor
        Me.CurrentX = OW
        Me.CurrentY = OH
        Me.Print STR
End Sub

Private Sub LOGOEXIT()
Dim I As Long
T.Enabled = False
U.Visible = False
Me.Cls
DrawBackground

Me.FontName = "Tahoma"
Me.FontSize = 55
Me.ForeColor = RGB(255, 255, 255)
Me.FontBold = True


For I = 0 To 9
    Me.CurrentX = 14
    Me.CurrentY = 60
    TEXTShadow Mid("UMAIR_11D", 1, CByte(I)), RGB(229, 237, 247)
    DoEvents
    Call Sleep(30)
Next
Me.CurrentX = Me.ScaleWidth - 25
Me.CurrentY = 70
Me.FontSize = 12
Me.Print "®"
Me.Refresh

Sleep 1000

Me.FontSize = 20
Me.ForeColor = RGB(255, 255, 255)
For I = 0 To 22
    Me.CurrentX = 53
    Me.CurrentY = 145
    TEXTShadow Mid("BRING FUTURE FEATURES", 1, CByte(I)), RGB(190, 210, 234)
    DoEvents
    Call Sleep(30)
Next

Sleep 1000
Me.FontSize = 10

For I = 0 To 92
    Me.CurrentX = 53
    Me.CurrentY = 165
    Me.Print Mid(".......................  ............................  ...................................", 1, CByte(I))
    DoEvents
    Call Sleep(20)
Next

Sleep 2000

Unload Me
End Sub

Private Sub T_Timer()
If U.Value = U.Max Then
LOGOEXIT
Else
U.Value = U.Value + 1
End If

End Sub

Private Sub TT_Timer()
LOGO
End Sub

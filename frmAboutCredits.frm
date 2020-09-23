VERSION 5.00
Begin VB.Form frmAboutCredits 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Your Program...."
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   368
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
   Begin VB.Timer ReDrawTimer 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   4800
      Top             =   480
   End
   Begin VB.PictureBox picBuffer 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   5265
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   5325
   End
   Begin VB.Image imgIcon 
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2000 "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblVersion"
      Height          =   195
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   675
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   360
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblTitle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "frmAboutCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Original Code by Mark Robert Strange (contact info unknown)
'Taken from Planet Source Code (www.planetsourcecode.com) on July 18, 2000
'Severly modified/cleaned up, etc by Duane Newman (duane.newman@dsionline.com)

Private Declare Function BitBlt Lib "gdi32" ( _
   ByVal hdcDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hDCSrc As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) _
   As Long

Private Const SRCCOPY = &HCC0020

Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type DisplayFormatCodes
    bBold As Boolean
    bItalic As Boolean
    iFont As Integer
    acAlignment As AlignmentConstants
    bPause As Boolean
    bPreviosLine As Boolean
End Type

Private CreditText() As String
Private CreditCode() As DisplayFormatCodes
Private NumLines As Integer

Private lX As Long
Private lY As Long

Private Function GetRESText(sResType As String, iResID As Integer) As String
                
Dim sResData As String
Dim vResData As Variant
Dim iLoop As Integer

    vResData = LoadResData(iResID, sResType)
    For iLoop = 0 To UBound(vResData)
        sResData = sResData & Chr(vResData(iLoop))
    Next iLoop

    GetRESText = sResData

End Function

Public Function ShowAboutSplash(Optional sResType As String = "NONE", Optional iResID As Integer)
    
    Me.Caption = ""
    
    'Set buffers to same size as output
    picBuffer.Left = 2
    picBuffer.Top = 2
    picBuffer.Width = Me.ScaleWidth - 1
    picBuffer.Height = Me.ScaleHeight - 1
    
    cmdOk.Visible = False
    lblTitle.Visible = False
    lblVersion.Visible = False
    Label1.Visible = False
    Line1.Visible = False
    imgIcon.Visible = False
    
    If sResType = "NONE" Then
        BuildCreditArray GetCreditTextFromFile
    Else
        BuildCreditArray GetRESText(sResType, iResID)
    End If

    lX = picBuffer.ScaleLeft
    lY = picBuffer.ScaleHeight
    ' Activate timer and start scrolling..
    ReDrawTimer.Enabled = True

    Me.Show

End Function

Public Function ShowAboutCredits(iconSource As Form, Optional sResType As String = "NONE", Optional iResID As Integer, Optional fscModal As FormShowConstants = vbModal)

    Me.Caption = "About " & App.Title
    Me.Icon = iconSource.Icon
    imgIcon.Picture = iconSource.Icon
    
    lblTitle = App.Title
    lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblVersion.Left = lblTitle.Left + lblTitle.Width - lblVersion.Width
    Label1.Caption = Label1.Caption & App.CompanyName
    
    If sResType = "NONE" Then
        BuildCreditArray GetCreditTextFromFile
    Else
        BuildCreditArray GetRESText(sResType, iResID)
    End If
    
    lX = picBuffer.ScaleLeft
    lY = picBuffer.ScaleHeight
    ' Activate timer and start scrolling..
    ReDrawTimer.Enabled = True
    
    Me.Show fscModal

End Function

Private Sub Form_Load()

    Me.ScaleMode = vbPixels
    
    'Setup Buffer for drawing
    picBuffer.ScaleMode = vbPixels

    picBuffer.ForeColor = vbWhite
    picBuffer.BackColor = vbBlack
    picBuffer.AutoRedraw = True

    picBuffer.Visible = False
            
End Sub
Private Function GetCreditTextFromFile() As String

    Dim sTemp As String
    
    Open (App.Path & "\credits.txt") For Input As #1
    
    sTemp = Input(LOF(1), 1)
    
    Close #1
    
    GetCreditTextFromFile = sTemp
    
End Function


Private Function BuildCreditArray(sTextIn As String)

Dim iCRLFPos As Integer
Dim iLastPos As Integer

    NumLines = -1
    
    iLastPos = 1
    iCRLFPos = InStr(iLastPos, sTextIn, vbCrLf)
    
    Do Until iCRLFPos = 0
        NumLines = NumLines + 1
        ReDim Preserve CreditText(0 To NumLines)
        ReDim Preserve CreditCode(0 To NumLines)
        CreditText(NumLines) = Mid$(sTextIn, iLastPos, iCRLFPos - iLastPos)
        
        ' Set Codes
        If StripCode(CreditText(NumLines), "<AL>") = True Then
            CreditCode(NumLines).acAlignment = vbLeftJustify
        ElseIf StripCode(CreditText(NumLines), "<AC>") = True Then
            CreditCode(NumLines).acAlignment = vbCenter
        ElseIf StripCode(CreditText(NumLines), "<AR>") = True Then
            CreditCode(NumLines).acAlignment = vbRightJustify
        Else
            If NumLines > 0 Then
                CreditCode(NumLines).acAlignment = CreditCode(NumLines - 1).acAlignment
            Else
                CreditCode(NumLines).acAlignment = vbCenter
            End If
        End If
        
        If NumLines > 0 Then
            CreditCode(NumLines).bBold = IIf(StripCode(CreditText(NumLines), "<B>"), True, IIf(StripCode(CreditText(NumLines), "</B>"), False, IIf(NumLines > 0, CreditCode(NumLines - 1).bBold, True)))
        Else
            CreditCode(NumLines).bBold = IIf(StripCode(CreditText(NumLines), "<B>"), True, IIf(StripCode(CreditText(NumLines), "</B>"), False, True))
        End If
        If NumLines > 0 Then
            CreditCode(NumLines).bItalic = IIf(StripCode(CreditText(NumLines), "<I>"), True, IIf(StripCode(CreditText(NumLines), "</I>"), False, IIf(NumLines > 0, CreditCode(NumLines - 1).bItalic, True))) 'IIf(StripCode(CreditText(NumLines), "<I>"), True, Not (StripCode(CreditText(NumLines), "</I>"))) 'StripCode(CreditText(NumLines), "<I>")
        Else
            CreditCode(NumLines).bItalic = IIf(StripCode(CreditText(NumLines), "<I>"), True, IIf(StripCode(CreditText(NumLines), "</I>"), False, True))
        End If
        CreditCode(NumLines).bPause = StripCode(CreditText(NumLines), "<P>")
        CreditCode(NumLines).bPreviosLine = StripCode(CreditText(NumLines), "<PL>")
        If InStr(CreditText(NumLines), "<F") > 0 Then
            CreditCode(NumLines).iFont = CInt(Mid$(CreditText(NumLines), InStr(CreditText(NumLines), "<F") + 2, 2))
            CreditText(NumLines) = Left$(CreditText(NumLines), InStr(CreditText(NumLines), "<F") - 1) & Mid$(CreditText(NumLines), InStr(CreditText(NumLines), "<F") + 5)
        Else
            If NumLines > 0 Then
                CreditCode(NumLines).iFont = CreditCode(NumLines - 1).iFont
            Else
                CreditCode(NumLines).iFont = 12
            End If
        End If
        
        'Get primed for next run
        iLastPos = iCRLFPos + 2
        iCRLFPos = InStr(iLastPos, sTextIn, vbCrLf)
        ' Check here for last line not ending in CRLF
        If iCRLFPos = 0 And iLastPos < Len(sTextIn) Then
            iCRLFPos = Len(sTextIn)
        End If
    
    Loop
'    Stop
End Function
Private Function StripCode(sIn As String, sCode As String) As Boolean

StripCode = False

Do While InStr(sIn, sCode) > 0
    sIn = Left$(sIn, InStr(sIn, sCode) - 1) & Mid$(sIn, InStr(sIn, sCode) + Len(sCode))
    StripCode = True
Loop


End Function

Private Sub RedrawTimer_Timer()

Dim l As Long
Dim j As Long
Dim iLineOffset As Integer
Dim r As RECT

On Error Resume Next
        
    ' Reset our time incase of Pause
    ReDrawTimer.Interval = 30
    
    ' Clear the Buffer for this run
    picBuffer.BackColor = vbBlack
    picBuffer.Cls
    
    ' Do the following for each line of text in our credits message...
    For j = 0 To NumLines Step 1
        
        ' Enter the CreditCode array..
        ' Set any line options like bold, italic, font size, etc..
        picBuffer.FontBold = CreditCode(j).bBold
        picBuffer.FontItalic = CreditCode(j).bItalic
        picBuffer.FontSize = CreditCode(j).iFont
        
        ' Set the starting location of where to print the text. Starts off below the bottom of the buffer.
        ' Don't increment the line if it is  a previous line draw
        If CreditCode(j).bPreviosLine = False Then
            iLineOffset = iLineOffset + (picBuffer.FontSize + 6)
        End If
        
        ' Move our CurrentX with the offset
        picBuffer.CurrentY = lY + iLineOffset 'lY + (j * picBuffer.FontSize + (6 * j))
        
        ' Select the alignment right here
        Select Case CreditCode(j).acAlignment
            Case vbLeftJustify
                picBuffer.CurrentX = 5
            Case vbCenter
                picBuffer.CurrentX = (picBuffer.ScaleWidth / 2) - (picBuffer.TextWidth(CreditText(j)) / 2)
            Case vbRightJustify
                picBuffer.CurrentX = picBuffer.ScaleWidth - picBuffer.TextWidth(CreditText(j)) - 5
        End Select

        ' Process Pause here.. If pause and line is centered then set timer to 1000/resets at begining of sub
        If CreditCode(j).bPause = True And picBuffer.CurrentY = CInt((picBuffer.ScaleHeight / 2) - (picBuffer.TextHeight(CreditText(j)) / 2)) Then ReDrawTimer.Interval = 1000
        
        ' Set ForeColor based on place in screen..
        ' TODO: Create var to set Color and break out RGB below
        picBuffer.ForeColor = RGB(FitRGB(100 - Abs(CurrentYColor)), FitRGB(255 - Abs(CurrentYColor)), FitRGB(100 - Abs(CurrentYColor)))
        
        If j = NumLines Then
            If picBuffer.CurrentY < -(picBuffer.TextHeight(CreditText(j)) + 10) Then
                ' If we've painted the last line, and it's above the top, there's no more text to scroll
                ' so now we run it again.. with just a little lead time so it looks a little better <DN>
                lY = picBuffer.ScaleHeight + 10
            End If
        End If
        
        ' Send the text directly into the buffer hDC
        picBuffer.Print CreditText(j)
        
    Next
    
    ' Ok, now that we have painted the entire buffer as we see fit for this pass, we blast the entire
    ' finished image directly to our output picturebox control.
    ' Now it goes straight to the form!
    l = BitBlt(Me.hDC, picBuffer.Left, picBuffer.Top, picBuffer.ScaleWidth, picBuffer.ScaleHeight, picBuffer.hDC, 0, 0, SRCCOPY)
        
    'Now we create a RECT for teh buffer size and possition
    r.Left = picBuffer.Left
    r.Top = picBuffer.Top
    r.Right = r.Left + picBuffer.ScaleWidth
    r.Bottom = r.Top + picBuffer.ScaleHeight
    
    ' We use that RECT to invalidate the area forcing the smooth refresh
    InvalidateRect Me.hwnd, r, 0
    
    ' Change the offset for the location of where the text will display next turn
    lY = lY - 1

End Sub
Private Function CurrentYColor()

' This code basically fades the text in from the bottom 10 or so visible pixels
' And out at the top 10 pixels, and goes full strength in the center

Dim myY As Integer
Dim myH As Integer

    myY = picBuffer.CurrentY
    myH = picBuffer.Height
    
    If myY < 10 Then
        CurrentYColor = Abs(myY - 10) / (20 / 255)
    ElseIf myY > (myH - 30) Then
        CurrentYColor = IIf(myY - (myH - 30) > 30, 30, myY - (myH - 30)) / (20 / 255)
    Else
        CurrentYColor = 0
    End If
    
    ' Uncomment this line and it will be full color in the center.
    ' It will fade in from the bottom to the center and back out to the top.
    'CurrentYColor = ((myH / 2) - myY) / ((myH / 2) / 255)

End Function
Private Function FitRGB(InVal As Long)

    FitRGB = IIf(InVal > 255, 255, IIf(InVal < 0, 0, InVal))

End Function

Private Sub cmdOk_Click()
    ReDrawTimer.Enabled = False
    Unload Me
End Sub



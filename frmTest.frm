VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTest 
   Caption         =   "Translucifer - Move Window To See More"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   182
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   387
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraContainer 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2670
      Left            =   3180
      TabIndex        =   6
      Top             =   0
      Width           =   2565
      Begin VB.TextBox txtVal 
         Enabled         =   0   'False
         Height          =   240
         Left            =   1950
         TabIndex        =   15
         Top             =   1365
         Width           =   525
      End
      Begin VB.PictureBox picColor 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   195
         Left            =   135
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   12
         Top             =   840
         Width           =   225
      End
      Begin VB.HScrollBar hsOpacity 
         Height          =   240
         Left            =   135
         Max             =   255
         TabIndex        =   11
         Top             =   1365
         Width           =   1770
      End
      Begin VB.OptionButton optTranslucency 
         Caption         =   "Translucency With Color"
         Height          =   375
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Top             =   450
         Width           =   2175
      End
      Begin VB.OptionButton optTranslucency 
         Caption         =   "Translucency With Picture"
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   180
         Width           =   2355
      End
      Begin VB.CommandButton cmdBail 
         Caption         =   "Exit"
         Height          =   360
         Left            =   1575
         TabIndex        =   7
         Top             =   2265
         Width           =   915
      End
      Begin MSComDlg.CommonDialog cdlgColor 
         Left            =   90
         Top             =   2100
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblInfo 
         Caption         =   "With this technique, less opacity means less flicker...."
         Height          =   420
         Left            =   135
         TabIndex        =   14
         Top             =   1650
         Width           =   2295
      End
      Begin VB.Label lblColor 
         Caption         =   "Click Box To Choose Color"
         Enabled         =   0   'False
         Height          =   225
         Left            =   435
         TabIndex        =   13
         Top             =   855
         Width           =   1995
      End
      Begin VB.Label lblOpacity 
         Caption         =   "Set Opacity:"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   945
      End
   End
   Begin VB.PictureBox picBB1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2580
      Index           =   2
      Left            =   6135
      Picture         =   "frmTest.frx":08CA
      ScaleHeight     =   172
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   203
      TabIndex        =   5
      Top             =   6165
      Width           =   3045
   End
   Begin VB.PictureBox picBB1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2580
      Index           =   1
      Left            =   3090
      Picture         =   "frmTest.frx":169E
      ScaleHeight     =   172
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   203
      TabIndex        =   4
      Top             =   6165
      Width           =   3045
   End
   Begin VB.PictureBox picBB1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2580
      Index           =   0
      Left            =   45
      Picture         =   "frmTest.frx":22E7
      ScaleHeight     =   172
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   203
      TabIndex        =   3
      Top             =   6165
      Width           =   3045
   End
   Begin VB.PictureBox picBlend 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2580
      Left            =   3285
      ScaleHeight     =   172
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   203
      TabIndex        =   2
      Top             =   3465
      Width           =   3045
   End
   Begin VB.PictureBox picBlit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   120
      ScaleHeight     =   172
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   203
      TabIndex        =   1
      Top             =   3465
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.PictureBox picDst 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   75
      ScaleHeight     =   172
      ScaleMode       =   0  'User
      ScaleWidth      =   203
      TabIndex        =   0
      Top             =   90
      Width           =   3075
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
'  Translucifer Project - My attempt at a "Pseudo" or "Poor Man's" translucency.
'  Uses Paul Caton's excellent Subclass and Timer classes which  can be found here:
'  http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=51403&lngWId=1
'  I highly recommend them for what I would characterize as 99.7% dummy-proof use.
'  This code works best compiled with all advanced optimizations checked...tested it
'  on Win2K.  Should work in Win98, WinME, Win2K, and WinXP.  How this looks will
'  vary from monitor to monitor.  I am hoping someone can help me eliminate or, at
'  least, minimize the flickering to a point where could be used in a production
'  application....anyone?
'
'  Copyright © 2004, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
Option Explicit

'**************************************************************************************************
' Translucifer App Constant Declares
'**************************************************************************************************
Private Const RGN_AND = &H1&
Private Const RGN_DIFF = &H4&
Private Const INIT_VAL = 64

'**************************************************************************************************
' Translucifer App Struct Declares
'**************************************************************************************************
Private Type BLENDFUNCTION
     BlendOp As Byte
     BlendFlags As Byte
     SourceConstantAlpha As Byte
     AlphaFormat As Byte
End Type ' BLENDFUNCTION

Private Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type ' RECT

'**************************************************************************************************
' Translucifer App API Declarations
'**************************************************************************************************
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, _
     ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
     ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, _
     ByVal dreamAKA As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
     ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
     ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, _
     ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, _
     ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, _
     ByVal y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
       ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
       
'**************************************************************************************************
' Translucifer App Implements
'**************************************************************************************************
Implements WinSubHook2.iSubclass
Implements WinSubHook2.iTimer

'**************************************************************************************************
' Translucifer Module-Level Variables
'**************************************************************************************************
Dim m_iCtr As Integer
Dim m_OpacityVal As Long
Dim m_SC As cSubClass
Dim m_Tmr As cTimer

'**************************************************************************************************
' Translucifer App VB-Intrinsic Events
'**************************************************************************************************
Private Sub cmdBail_Click()
     ' Bail
     Unload Me
End Sub ' cmdBail_Click

Private Sub Form_Load()
     ' Draw the initial Translucifer
     Translucify
     ' Instantiate subclass object
     Set m_SC = New cSubClass
     ' Begin subclassing
     Call m_SC.Subclass(Me.hwnd, Me)
     ' Add the messages we care about.  We want to detect
     ' when the form is moved or activated....
     m_SC.AddMsg WM_MOVE, MSG_AFTER
     m_SC.AddMsg WM_ACTIVATE, MSG_AFTER
     m_SC.AddMsg WM_ACTIVATEAPP, MSG_BEFORE
     ' Instantiate timer object
     Set m_Tmr = New cTimer
     ' Start timer and set interval
     m_Tmr.TmrStart Me, 150
     ' set option for picture Translucifer
     optTranslucency(0) = True
     ' Set initial opacity level
     hsOpacity.Value = INIT_VAL
     ' set textbox to value
     txtVal = CStr(INIT_VAL)
End Sub ' Form_Load

Private Sub Form_Unload(Cancel As Integer)
     ' Kill the timer
     m_Tmr.TmrStop
     ' Destroy timer object
     Set m_Tmr = Nothing
     'Kill the subclasser
     m_SC.UnSubclass
     ' Destroy Subclass object
     Set m_SC = Nothing
     ' Habit
     Set frmTest = Nothing
End Sub ' Form_Unload

Private Sub hsOpacity_Scroll()
     m_OpacityVal = hsOpacity.Value
     txtVal = CStr(m_OpacityVal)
End Sub ' hsOpacity_Scroll

Private Sub optTranslucency_Click(Index As Integer)
     Select Case Index
          Case 0
               ' Start timer and set interval
               m_Tmr.TmrStart Me, 150
               ' disable color picker
               picColor.Enabled = False
               lblColor.Enabled = False
               ' draw
               Translucify
          Case 1
               ' Kill the timer
               m_Tmr.TmrStop
               ' lose the picture in picBlend
               picBlend = Nothing
               ' clear picture
               picBlend.Cls
               ' set color
               picBlend.BackColor = picColor.BackColor
               ' enable color picker
               picColor.Enabled = True
               ' Draw
               Translucify
     End Select
End Sub ' optTranslucency_Click

Private Sub picColor_Click()
     Dim lColor As Long
     ' set flags
     cdlgColor.flags = cdlCCRGBInit
     ' show dialog
     cdlgColor.ShowColor
     ' Set color
     picColor.BackColor = cdlgColor.Color
     picBlend.BackColor = cdlgColor.Color
     ' draw
     Translucify
End Sub ' picColor_Click

'**************************************************************************************************
' Paul Caton's Righteous Subclasser & Timer Procs
'**************************************************************************************************
Private Sub iSubclass_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As WinSubHook2.eMsg, wParam As Long, lParam As Long)
     Select Case uMsg
          Case WM_ACTIVATE, WM_ACTIVATEAPP, WM_MOVE
               ' call translucent method
               Translucify
          Case Else
               ' we'll think of something later
     End Select
End Sub ' iSubclass_Proc

Private Sub iTimer_Proc(ByVal lElapsedMS As Long, ByVal lTimerID As Long)
     ' increment counter
     m_iCtr = m_iCtr + 1
     ' set blend picturebox to item number of b&b pic array
     picBlend = picBB1(m_iCtr - 1)
     ' reset counter so we don't go out of range
     If m_iCtr = 3 Then m_iCtr = 0
End Sub ' iTimer_Proc

'**************************************************************************************************
' Translucifer App Custom Methods
'**************************************************************************************************
Private Sub Translucify()
     Dim hR1 As Long
     Dim hR2 As Long
     Dim hR3 As Long
     Dim hR4 As Long
     Dim lhRgn As Long
     Dim lRtn As Long
     Dim rc As RECT
     Dim lBlend As Long
     Dim bf As BLENDFUNCTION
     ' Create the initial region encompassing our client window
     rc.Left = 0
     rc.Top = -((frmTest.Height - frmTest.ScaleHeight) \ Screen.TwipsPerPixelY)
     rc.Right = frmTest.Width
     rc.Bottom = frmTest.Height
     ' create region
     hR1 = CreateRectRgn(rc.Left, rc.Top, rc.Right, rc.Bottom)
     ' Create region where we want to "make a rectangle hole."  You could
     ' use any shapes region here but for this app you should leave this
     ' be so you won't have to make a ton of adjustments.
     rc.Left = 10
     rc.Top = 30
     rc.Right = 213
     rc.Bottom = 200
     ' Create our rectangle region
     hR2 = CreateRectRgn(rc.Left, rc.Top, rc.Right, rc.Bottom)
     ' create an empty region to combine the regions we created.
     lhRgn = CreateRectRgn(0, 0, 0, 0)
     ' combine our regions
     hR3 = CombineRgn(lhRgn, hR1, hR2, RGN_DIFF)
     ' Delete our other temporary regions
     DeleteObject hR1
     DeleteObject hR2
     DeleteObject hR3
     ' Set the final region on our parent form
     SetWindowRgn frmTest.hwnd, lhRgn, True
     ' Give our region a chance to be created before getting a
     ' snapshot of it.
     DoEvents
     ' Now get a snapshot of the region
     lRtn = BitBlt(picBlit.hDC, 0, 0, 203, 172, frmTest.hDC, 6, 7, vbSrcCopy)
     ' alphablend the picture/color to the screenshot we took of the region
     ' into the destination picturebox.
     DoEvents
     ' Draw the background picture with full opacity
     bf.BlendOp = AC_SRC_OVER
     bf.BlendFlags = 0
     bf.SourceConstantAlpha = 255
     bf.AlphaFormat = 0
     CopyMemory lBlend, bf, 4
     AlphaBlend picDst.hDC, 0, 0, picDst.ScaleWidth, picDst.ScaleHeight, _
          picBlit.hDC, 0, 0, picBlit.ScaleWidth, picBlit.ScaleHeight, lBlend
     ' Now draw the second picture with variable transparency over the first picture
     bf.SourceConstantAlpha = IIf(m_OpacityVal, m_OpacityVal, INIT_VAL)
     CopyMemory lBlend, bf, 4
     AlphaBlend picDst.hDC, 0, 0, picDst.ScaleWidth, picDst.ScaleHeight, _
          picBlend.hDC, 0, 0, picBlend.ScaleWidth, picBlend.ScaleHeight, lBlend
     picDst.Refresh
     ' lose the region
     SetWindowRgn frmTest.hwnd, 0, True
End Sub ' Translucify



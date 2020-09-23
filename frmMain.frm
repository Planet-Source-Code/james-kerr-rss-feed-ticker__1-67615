VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "RSS Feed Reader"
   ClientHeight    =   2415
   ClientLeft      =   1515
   ClientTop       =   2190
   ClientWidth     =   7455
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   7455
   Begin VB.CheckBox chkLong 
      Caption         =   "Use long item descriptions"
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
      Left            =   2610
      TabIndex        =   9
      Top             =   1125
      Width           =   4560
   End
   Begin VB.TextBox txtSpeed 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1575
      TabIndex        =   8
      Text            =   "10"
      Top             =   1080
      Width           =   600
   End
   Begin VB.TextBox txtURL 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1575
      TabIndex        =   6
      Top             =   630
      Width           =   5640
   End
   Begin VB.TextBox txtFeed 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1575
      TabIndex        =   4
      Top             =   180
      Width           =   5640
   End
   Begin VB.PictureBox picMessageIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   720
      Picture         =   "frmMain.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   6210
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picMessage 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1035
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   130
      TabIndex        =   1
      Top             =   6210
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.PictureBox picTicker 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   990
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   0
      Top             =   5805
      Width           =   1965
   End
   Begin VB.Timer timMessage 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   225
      Top             =   6120
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ticker Speed:"
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
      Index           =   2
      Left            =   180
      TabIndex        =   7
      Top             =   1125
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RSS Feed URL:"
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
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   675
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RSS Feed Title:"
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
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   225
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   3645
      Y1              =   5085
      Y2              =   5085
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
' RSS News Feed Ticker Example
' for Visual Basic 6
'
' This simple example shows you how to implement an RSS feed-based
' 'Ticker' in your VB Applications.
'
' Feel free to use/amend this code for your own applications.
' This code is provided 'as is' and we take no resposibility for
' any errors or omissions.
'
' ©2006 James Kerr, iSYS Systems Integration Ltd.
' Web: www.callmap.co.uk
'****************************************************************

Option Explicit
'Declarations
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                            ByVal X As Long, _
                                            ByVal Y As Long, _
                                            ByVal nWidth As Long, _
                                            ByVal nHeight As Long, _
                                            ByVal hSrcDC As Long, _
                                            ByVal xSrc As Long, _
                                            ByVal ySrc As Long, _
                                            ByVal dwRop As Long) As Long
Private Declare Function InternetGetConnectedState Lib "wininet" (lpdwFlags As Long, ByVal dwReserved As Long) As Boolean
Private iTickerSpeed As Integer
Private lTickerPos As Long
Private bShort As Boolean

Private Sub chkLong_Click()
    If chkLong.Value = 1 Then
        bShort = False
    Else
        bShort = True
    End If
    Call pvTickerMessage
End Sub

Private Sub Form_Load()
    Dim lFlags As Long
    Dim bRet As Boolean
    Dim ctlItem As Control
    'Check that we're connected to the internet
    
    bRet = InternetGetConnectedState(lFlags, 0)
    If bRet = False Then 'not connected - don'start the ticker!
        MsgBox "You're not connected to the internet!", vbExclamation
        For Each ctlItem In Me.Controls 'Disable input
            Select Case TypeName(ctlItem)
                Case "TextBox", "CheckBox", "Label"
                    ctlItem.Enabled = False
            End Select
        Next
        Exit Sub
    End If
    
    'set the initial values for the ticker
    iTickerSpeed = 10
    bShort = True
    'change these values to display your favourite ticker ....
    txtFeed.Text = "BBC News UK RSS Feed"
    txtURL.Text = "http://newsrss.bbc.co.uk/rss/newsonline_uk_edition/front_page/rss.xml"
    
    'call the ticker initialisation routine
    Call pvTickerMessage
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Line1.X2 = Me.ScaleWidth
    Line1.Y1 = Me.ScaleHeight - (picTicker.Height + 45)
    Line1.Y2 = Line1.Y1
    picTicker.Move 0, Me.ScaleHeight - picTicker.Height, Me.ScaleWidth
    picMessage.Height = (picMessage.TextHeight("A") * Screen.TwipsPerPixelY) + 20
    picTicker.Height = picMessage.Height
End Sub

Private Sub picTicker_DblClick()
    MsgBox "Shell Execute URL here", vbInformation
End Sub

Private Sub pvTickerMessage()

    Dim MyMessage As String
    Dim lRet As Long
    Dim lPos As Long
    Dim RSSChannel As Object
    Dim RSSItem As Object
    Dim objNode As Object
    
    On Error GoTo NoFeed
    picTicker.Cls 'clear the message panel
    MyMessage = txtFeed & ":" 'put the RSS title on the message
    Set RSSChannel = CreateObject("MSXML2.DOMDocument") 'create the XML object to hold the RSS channel
    RSSChannel.async = False 'wait till the data's returned
    RSSChannel.Load txtURL.Text 'load the 'channel' into the XML document
    Set RSSItem = RSSChannel.getElementsByTagName("item") 'get a list of <item> elements
    If Not RSSItem Is Nothing Then
        For lPos = 0 To (RSSItem.length - 1) 'for each <item> we find
            If bShort Then
                Set objNode = RSSItem(lPos).getElementsByTagName("title") 'Get the <title> element
            Else
                Set objNode = RSSItem(lPos).getElementsByTagName("description") 'Get the <description> element
            End If
            'note you could also get the URL element to associate it with a section of the message ...
            '(You'd need yo 'break up the message display to do this), but it would allow for
            'clicking on a specific item and loading the related page.
            MyMessage = MyMessage & " ...  " & objNode(0).Text & " " 'and add the text to the message
            Set objNode = RSSItem.nextNode 'get the next node in the element
        Next
    End If
NoFeed:
    On Error Resume Next
    MyMessage = MyMessage & String((iTickerSpeed / 2), " ") 'add some trailing space
    lPos = ((picMessage.Height / Screen.TwipsPerPixelY) - 16) / 2
    picMessage.Cls
    picMessage.Width = ((picMessage.TextWidth(MyMessage)) * Screen.TwipsPerPixelX) + 270
    lRet = BitBlt(picMessage.hDC, 0, lPos, 16, 16, picMessageIcon.hDC, 0, 0, vbSrcCopy) 'put the RSS Icon onto the message
    lTickerPos = Me.ScaleWidth / Screen.TwipsPerPixelX 'start at the right, and move to the left
    picMessage.CurrentX = 16
    picMessage.CurrentY = 0
    picMessage.Print MyMessage 'and put the message onto the picture box
    If Not timMessage.Enabled Then timMessage.Enabled = True 'initialise the ticker timer
    
End Sub

Private Sub timMessage_Timer()
    Dim lRet As Long
    On Error Resume Next
    lRet = BitBlt(picTicker.hDC, lTickerPos, 0, picMessage.ScaleWidth, picMessage.ScaleHeight, picMessage.hDC, 0, 0, vbSrcCopy)
    lTickerPos = lTickerPos - iTickerSpeed 'the amount we move the ticker by - higher value is faster
    If lTickerPos < (0 - picMessage.ScaleWidth) Then 'we've completed the message - reset and start again
        Call pvTickerMessage
    End If
End Sub

Private Sub txtFeed_Change()
    'You might want to set up a routine that refreshed the ticker with your new
    'information here (see the 'txtSpeed_Change' sub)
End Sub

Private Sub txtSpeed_Change()
    If IsNumeric(txtSpeed.Text) Then
        iTickerSpeed = CInt(txtSpeed.Text)
        Call pvTickerMessage
    End If
End Sub

Private Sub txtURL_Change()
    'You might want to set up a routine that refreshed the ticker with your new
    'information here (see the 'txtSpeed_Change' sub)
End Sub

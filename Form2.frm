VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form21 
   BorderStyle     =   0  'None
   Caption         =   "iPage"
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   3780
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2115
      ScaleWidth      =   4875
      TabIndex        =   3
      Top             =   480
      Width           =   4935
      Begin VB.VScrollBar VScroll1 
         Height          =   2120
         LargeChange     =   8
         Left            =   4620
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   2.45745e5
         Left            =   60
         ScaleHeight     =   2.45745e5
         ScaleWidth      =   4995
         TabIndex        =   4
         Top             =   0
         Width           =   5000
         Begin VB.Label Label3 
            Caption         =   "yes"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   1800
            Visible         =   0   'False
            Width           =   15
         End
         Begin VB.Label lblUnder 
            Caption         =   "False"
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   1800
            Visible         =   0   'False
            Width           =   15
         End
         Begin VB.Label lblItalic 
            Caption         =   "False"
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   1800
            Visible         =   0   'False
            Width           =   15
         End
         Begin VB.Label lblBold 
            Caption         =   "False"
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   1800
            Visible         =   0   'False
            Width           =   15
         End
         Begin VB.Label Label2 
            Caption         =   "0"
            Height          =   255
            Left            =   -120
            TabIndex        =   8
            Top             =   1800
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   1800
            Visible         =   0   'False
            Width           =   15
         End
      End
   End
   Begin MSWinsockLib.Winsock wData 
      Left            =   840
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   3160
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   4335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1A94
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1C70
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":2028
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":2204
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":23E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":25BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":2798
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":2974
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":2B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":2D2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "under"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "elite"
            Object.ToolTipText     =   "Elited Text"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "hacker"
            Object.ToolTipText     =   "Hacker Coded"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "back"
            Object.ToolTipText     =   "Backwards Text"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "echo"
            Object.ToolTipText     =   "Echoed Text"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "alter"
            Object.ToolTipText     =   "Alternative Caps"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "opt"
            Object.ToolTipText     =   "Show Options"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "conn"
            Object.ToolTipText     =   "Start Connection"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "dconn"
            Object.ToolTipText     =   "Disconnect"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   4800
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   5040
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   120
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   360
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Declarations for the Variables
Dim bolds As String
Dim italics As String
Dim unders As String
Dim R1 As Integer
Dim G1 As Integer
Dim B1 As Integer
Dim R2 As Integer
Dim G2 As Integer
Dim B2 As Integer
Dim sr1 As Integer
Dim sg1 As Integer
Dim sb1 As Integer
Dim sr2 As Integer
Dim sg2 As Integer
Dim sb2 As Integer
Dim messa As String
Dim nickn As String

'***************************'

Private Sub Command1_Click()
On Error Resume Next
' Dont Display the Message if the message is empty
If Text2.Text = "" Then
Exit Sub
End If
' Add an empty Character at the end of every message
Text2.Text = Text2.Text & " "
' Change the label caption to Yes so to display a nick name
Label1.Caption = "Yes"
' Filter The Language you can add more filtering if you want to
Text2.Text = ReplaceChars(Text2.Text, "ass", "***")
Text2.Text = ReplaceChars(Text2.Text, "fuck", "****")
Text2.Text = ReplaceChars(Text2.Text, "tits", "****")
Text2.Text = ReplaceChars(Text2.Text, "cock", "****")
' Call the Send Data Procedure
Call senddata
' Call The Command2_click Procedure
Call Command2_Click

End Sub

'***************************'

Private Sub Command2_Click()
' Check to see if label1 Caption is Yes if true then display the nickname
If Label1.Caption = "Yes" Then
Picture2.FontItalic = False
Picture2.FontUnderline = False
Picture2.ForeColor = vbBlue
Picture2.FontBold = True
Picture2.Print Form1.Text7.Text & "> ";
Label1.Caption = "dsd"
Picture2.FontBold = False
Else
End If
' Fade the Text in the picturebox
Fa Form1.Text3.Text, Form1.Text2.Text, Form1.Text1.Text, Form1.Text6.Text, Form1.Text5.Text, Form1.Text4.Text, Text2.Text
' Add the Times the Send Button Is Clicked
Label2.Caption = Add(Label2.Caption, 1)
' If the send button is clicked 2 times then scroll picture box
If Label2.Caption = 2 Then
VScroll1.Value = VScroll1.Value + 180
Label2.Caption = 0
Else
End If
' Empty Message when fading completes
Text2.Text = ""
End Sub

'***************************'

Private Sub Form_Load()
' Make the form transparent
SetRegion
' Set the protocol
wData.Protocol = sckUDPProtocol
End Sub

'***************************'

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Drag the form without title bar
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

'***************************'

Private Sub Image2_Click()
' End the Program if the user clicks the end button
End
End Sub

'***************************'

Private Sub Image3_Click()
' Minimize the program if user clicks the minimize button
Me.WindowState = 1
End Sub

'***************************'

Private Sub Image4_Click()
If Me.Width = 5550 Then
Me.Width = 5370
Me.Height = 360
ElseIf Me.Width = 5370 Then
Me.Width = 5550
Me.Height = 3780
End If
End Sub

'***************************'

Private Sub Text2_Change()
Command1.Default = True
End Sub

'***************************'

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "bold"
' Make the label caption to True if it is false and vice versa
If lblBold.Caption = "False" Then
lblBold.Caption = "True"
Else
lblBold.Caption = "False"
End If
Case "italic"
' Make the label caption to True if it is false and vice versa
If lblItalic.Caption = "False" Then
lblItalic.Caption = "True"
Else
lblItalic.Caption = "False"
End If
Case "under"
' Make the label caption to True if False and vice versa
If lblUnder.Caption = "False" Then
lblUnder.Caption = "True"
Else
lblUnder.Caption = "False"
End If
Case "elite"
' Make text Elited
Text2.Text = fn138((Text2.Text))
Case "hacker"
' Make Hacker Text
Text2.Text = fn170((Text2.Text))
Case "back"
' Reverse Text
Text2.Text = fn100((Text2.Text))
Case "alter"
' Make Alternative Caps
Text2.Text = fnC8((Text2.Text))
Case "echo"
' Echoed Text
Text2.Text = EchoText(Text2.Text, False)
Case "opt"
' Display Options
Form1.Show 1
Case "conn"
' Connect To Computer
Call connect
Case "dconn"
' Disconnect from remote computer
Call dconnect
End Select
End Sub

'***************************'

Function Fa(R1%, G1%, B1%, R2%, G2%, B2%, TheText$, Optional nickna$)
' This function will fade the text and display it in the picture box
On Error GoTo error
' Make sure the message storing text box is empty
Text1.Text = ""
' Display the remote nickname if wanted
If nickna = "" Then
Else
Picture2.ForeColor = vbRed
Picture2.FontUnderline = False
Picture2.FontItalic = False
Picture2.FontBold = True
Picture2.Print nickna & "> ";
Label1.Caption = "dsd"
Picture2.FontBold = False
End If
textlen$ = Len(TheText)
For i = 1 To textlen$
textdone$ = Left(TheText, i)
lastchr$ = Right(textdone$, 1)
colorX = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
' Store the no. of characters
Text1.Text = Text1.Text & lastchr$
' Change the forecolor
Picture2.ForeColor = colorX
' Make text bold if required
If lblBold.Caption = "True" Then
Picture2.FontBold = True
ElseIf lblBold.Caption = "False" Then
Picture2.FontBold = False
End If
' Make text Italic If Required
If lblItalic.Caption = "True" Then
Picture2.FontItalic = True
ElseIf lblItalic.Caption = "False" Then
Picture2.FontItalic = False
End If
' Make Text Underlined if required
If lblUnder.Caption = "True" Then
Picture2.FontUnderline = True
ElseIf lblUnder.Caption = "False" Then
Picture2.FontUnderline = False
End If
' If the characters reaches 40 mark then start with a new line
If Len(Text1.Text) > 40 Then
' Empty Storing text box
Text1.Text = ""
Picture2.Print lastchr$
GoTo jo
Else
End If
fig:
Dim gh As Long
If gh = 2 Then
gh = 0
Picture2.Print lastchr$
Text1.Text = ""
Exit Function
Else
End If
If Len(Text1.Text) = Len(Text2.Text) - 1 Then
gh = Add(gh, 1)
GoTo fig
Else
Picture2.Print lastchr$;
End If
jo:
Next i
Text1.Text = ""
Picture2.Print " "
Exit Function
error:
End Function

'***************************'

Private Sub VScroll1_Change()
' Scroll the picture box as required
If VScroll1.Value >= 32750 Then
Picture2.Cls
VScroll1.Value = 0
End If
Picture2.Top = "-" & VScroll1.Value
End Sub

'***************************'

Function Add(num1 As Long, num2 As Long) As Long
' Function to add two given numbers
    On Error Resume Next
    Add = Val(num1) + Val(num2)
End Function

'***************************'

Function Subtract(num1 As Long, num2 As Long) As Long
' Function to subtract two numbers
    On Error Resume Next
Subtract = Val(num1) - Val(num2)
End Function

'***************************'

Public Sub connect()
On Error Resume Next
' Close current connection if active
wData.Close
' Specify Details of the Remote Computer
' *****************'
' Set the remote ip
wData.RemoteHost = Form1.Text8.Text
' Set the local port
wData.LocalPort = Form1.Text9.Text
' Set the remote port
wData.RemotePort = Form1.Text9.Text
wData.Bind
' Display Message Window
message "Connecting to remote computer", "1000"
' Enable the Command Button
Command1.Enabled = True
If Err Then MsgBox Err.Description
End Sub

'***************************'

Private Sub wData_Close()
' On Closing Display the Closing Message
message "Closing Current Connection", "1000"
End Sub

'***************************'

Private Sub wData_Connect()
' On Connection Display the connected message
message "Connected.", "700"
End Sub

'***************************'

Private Sub wData_ConnectionRequest(ByVal requestID As Long)
    ' Close Current Connection
    wData.Close
    wData.Accept requestID
    ' Display Message
    message "Connection accepted", "600"
End Sub

'***************************'

Private Sub wData_DataArrival(ByVal bytesTotal As Long)
Dim ndata As String
On Error Resume Next
'Get the data being sent to us:
wData.GetData ndata
' Parse all the data
bolds = GetInBetween(ndata, "[b]", "[/b]")
italics = GetInBetween(ndata, "[i]", "[/i]")
unders = GetInBetween(ndata, "[u]", "[/u]")
R1 = GetInBetween1(ndata, "[r1]", "[/r1]")
G1 = GetInBetween1(ndata, "[g1]", "[/g1]")
B1 = GetInBetween1(ndata, "[b1]", "[/b1]")
R2 = GetInBetween1(ndata, "[r2]", "[/r2]")
G2 = GetInBetween1(ndata, "[g2]", "[/g2]")
B2 = GetInBetween1(ndata, "[b2]", "[/b2]")
nickn = GetInBetween(ndata, "[ni]", "[/ni]")
messa = GetInBetween(ndata, "[m]", "[/m]")
' Set all the data as given
lblBold.Caption = bolds
lblItalic.Caption = italics
lblUnder.Caption = unders
' Call the fading text procedure
Fa R1, G1, B1, R2, G2, B2, messa, nickn
If Err Then MsgBox Err.Description
End Sub

'***************************'

Private Sub wData_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
' On error display a message box with its description
MsgBox Description
End Sub

'***************************'

Public Function message(mess As String, interval As Integer)
' Function for displaying the message window
' ***********************'
' Set given message
Form3.Label1.Caption = mess
' Set the interval
Form3.Timer1.interval = interval
' Enable the timer
Form3.Timer1.Enabled = True
' Show the form
Form3.Show 1
End Function

'***************************'

Public Sub dconnect()
' Dsiconnect Remote Computer
wData.Close
' Set the enable property of send button to false
Command1.Enabled = False
End Sub

'***************************'

Public Sub senddata()
Dim strMess As String
' Set the string to the required data
strMess = "[b]" & lblBold.Caption & "[/b]" & "[i]" & lblItalic.Caption & "[/i]" & "[u]" & lblUnder.Caption & "[/u]" & "[r1]" & Form1.Text3.Text & "[/r1]" & "[g1]" & Form1.Text2.Text & "[/g1]" & "[b1]" & Form1.Text1.Text & "[/b]" & "[r2]" & Form1.Text6.Text & "[/r2]" & "[g2]" & Form1.Text5.Text & "[/g2]" & "[b2]" & Form1.Text4.Text & "[/b2]" & "[m]" & Text2.Text & "[/m]" & "[ni]" & Form1.Text7.Text & "[/ni]"
' Send the Data
wData.senddata strMess
End Sub

'***************************'

Private Sub SetRegion()
'Free the memory allocated by the previous Region
    If hRgn Then DeleteObject hRgn
'Scan the Bitmap and remove all transparent pixels from
'it, creating a new region
    hRgn = GetBitmapRegion(Form21.Picture, 16711935)
'Set the Forms new Region
    SetWindowRgn Form21.hwnd, hRgn, True
End Sub

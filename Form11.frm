VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change Fading Gradient"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1215
      Left            =   120
      TabIndex        =   24
      Top             =   3120
      Width           =   4335
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3120
         TabIndex        =   32
         Text            =   "0"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3120
         TabIndex        =   30
         Text            =   "1113"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   28
         Text            =   "0.0.0.0"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   26
         Text            =   "nick1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Local IP:"
         Height          =   255
         Left            =   2400
         TabIndex        =   31
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Port:"
         Height          =   255
         Left            =   2400
         TabIndex        =   29
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "RemoteHost:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nickname:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   4440
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tools"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.PictureBox Picture10 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   4035
         TabIndex        =   22
         Top             =   600
         Width           =   4095
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H00000000&
         Height          =   735
         Left            =   2280
         ScaleHeight     =   675
         ScaleWidth      =   1275
         TabIndex        =   21
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0"
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0"
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0"
         Top             =   2040
         Width           =   495
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00000000&
         Height          =   735
         Left            =   720
         ScaleHeight     =   675
         ScaleWidth      =   1275
         TabIndex        =   17
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0"
         Top             =   2040
         Width           =   495
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00FF0000&
         Height          =   135
         Left            =   2280
         ScaleHeight     =   75
         ScaleWidth      =   1875
         TabIndex        =   13
         Top             =   1830
         Width           =   1935
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H0000FF00&
         Height          =   135
         Left            =   2280
         ScaleHeight     =   75
         ScaleWidth      =   1875
         TabIndex        =   12
         Top             =   1470
         Width           =   1935
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   2280
         ScaleHeight     =   75
         ScaleWidth      =   1875
         TabIndex        =   11
         Top             =   1110
         Width           =   1935
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FF0000&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   1875
         TabIndex        =   10
         Top             =   1830
         Width           =   1935
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H0000FF00&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   1875
         TabIndex        =   9
         Top             =   1470
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   1875
         TabIndex        =   8
         Top             =   1110
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll6 
         Height          =   135
         Left            =   2280
         Max             =   255
         TabIndex        =   7
         Top             =   1680
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   135
         Left            =   2280
         Max             =   255
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   135
         Left            =   120
         Max             =   255
         TabIndex        =   5
         Top             =   1680
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   135
         Left            =   120
         Max             =   255
         TabIndex        =   4
         Top             =   1320
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   135
         Left            =   120
         Max             =   255
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Left            =   2280
         Max             =   255
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   4035
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
' Hide the form
Me.Hide
End Sub

Private Sub Form_Load()
' Display the local ip address
Text10.Text = Form21.wData.LocalIP
End Sub

Private Sub HScroll1_Change()
' Change the text box color value to new color value
Text4.Text = Str(HScroll1.Value)
' Change the Picturebox Backcolor to new color
Picture9.BackColor = RGB(Text4.Text, Text5.Text, Text6.Text)
' Call the update sub to display a new gradient in the preview picture box
update
End Sub

Private Sub HScroll2_Change()
' Change the text box color value to new color value
Text1.Text = Str(HScroll2.Value)
' Change the Picturebox Backcolor to new color
Picture8.BackColor = RGB(Text1.Text, Text2.Text, Text3.Text) 'asign a picture box the value of the text boxs, which is the value of the scroll bars,....and because the scroll bars have a max of 255..it is also the red,green,blue value
' Call the update sub to display a new gradient in the preview picture box
update
End Sub

Private Sub HScroll3_Change()
' Change the text box color value to new color value
Text2.Text = Str(HScroll3.Value)
' Change the Picturebox Backcolor to new color
Picture8.BackColor = RGB(Text1.Text, Text2.Text, Text3.Text) 'asign a picture box the value of the text boxs, which is the value of the scroll bars,....and because the scroll bars have a max of 255..it is also the red,green,blue value
' Call the update sub to display a new gradient in the preview picture box
update
End Sub

Private Sub HScroll4_Change()
' Change the text box color value to new color value
Text3.Text = Str(HScroll4.Value)
' Change the Picturebox Backcolor to new color
Picture8.BackColor = RGB(Text1.Text, Text2.Text, Text3.Text) 'asign a picture box the value of the text boxs, which is the value of the scroll bars,....and because the scroll bars have a max of 255..it is also the red,green,blue value
' Call the update sub to display a new gradient in the preview picture box
update
End Sub

Private Sub HScroll5_Change()
' Change the text box color value to new color value
Text5.Text = Str(HScroll5.Value)
' Change the Picturebox Backcolor to new color
Picture9.BackColor = RGB(Text4.Text, Text5.Text, Text6.Text) 'asign a picture box the value of the text boxs, which is the value of the scroll bars,....and because the scroll bars have a max of 255..it is also the red,green,blue value
' Call the update sub to display a new gradient in the preview picture box
update
End Sub

Private Sub HScroll6_Change()
' Change the text box color value to new color value
Text6.Text = Str(HScroll6.Value)
' Change the Picturebox Backcolor to new color
Picture9.BackColor = RGB(Text4.Text, Text5.Text, Text6.Text) 'asign a picture box the value of the text boxs, which is the value of the scroll bars,....and because the scroll bars have a max of 255..it is also the red,green,blue value
' Call the update sub to display a new gradient in the preview picture box
update
End Sub

Public Function GradEffects(FormGrad As PictureBox, FadeEffect As Integer, TimesToFade As Integer, colora As ColorConstants, colorb As ColorConstants, Optional colorc As ColorConstants, Optional colord As ColorConstants, Optional colore As ColorConstants, Optional colorf As ColorConstants, Optional colorg As ColorConstants, Optional colorh As ColorConstants, Optional colori As ColorConstants, Optional colorj As ColorConstants)

    Dim CurrentFadeNum As Integer
    Dim RR As Integer, GG As Integer, BB As Integer
    Dim rrr As Integer, ggg As Integer, bbb As Integer
    'Checks to make sure its a valid fade effect
    If FadeEffect > 17 Or FadeEffect < 1 Then MsgBox "Error: Invalid Fade Effect", vbCritical, "Error": Exit Function
    'clears the form


    FormGrad.Cls
        'Makes the coding easier
        TimesToFade = TimesToFade - 1
        'Line width 3 so no white spots
        'Increase line width if for some reason there are spots which are
        '     left undrawn


        FormGrad.DrawWidth = 3
            'the higher the number for the scalewidth
            'the better the quality but slower


            FormGrad.ScaleWidth = 255 * 2
                'AutoRedraw is needed so the form doesn't clear itself when it's
                '     not visible


                FormGrad.AutoRedraw = True
                    'Keep this line, makes sure form height and width are equal for e
                    '     asier coding


                    FormGrad.ScaleHeight = FormGrad.ScaleWidth
                        'Makes a loop for each time the fade changes
                        'CurrentFadeNum would be the current fade number


                        For CurrentFadeNum = 1 To TimesToFade
                            
                            'previous color it fades to is now the fading color
                            RR = rrr
                            GG = ggg
                            BB = bbb
                            
                            'Sets the RGB for each color
'You can add more for more color fades
'Thank you http://www.planet-source-code.com/ for this RGB code
Select Case CurrentFadeNum

Case 1
RR = colora Mod &H100
GG = (colora \ &H100) Mod &H100
BB = (colora \ &H10000) Mod &H100
rrr = colorb Mod &H100
ggg = (colorb \ &H100) Mod &H100
bbb = (colorb \ &H10000) Mod &H100
'Adds the background color of the first color when using fade eff
'     ect 10 or 11
If FadeEffect = 10 Or FadeEffect = 11 Then FormGrad.BackColor = RGB(RR, GG, BB)
Case 2
rrr = colorc Mod &H100
ggg = (colorc \ &H100) Mod &H100
bbb = (colorc \ &H10000) Mod &H100
Case 3
rrr = colord Mod &H100
ggg = (colord \ &H100) Mod &H100
bbb = (colord \ &H10000) Mod &H100
Case 4
rrr = colore Mod &H100
ggg = (colore \ &H100) Mod &H100
bbb = (colore \ &H10000) Mod &H100
Case 5
rrr = colorf Mod &H100
ggg = (colorf \ &H100) Mod &H100
bbb = (colorf \ &H10000) Mod &H100
Case 6
rrr = colorg Mod &H100
ggg = (colorg \ &H100) Mod &H100
bbb = (colorg \ &H10000) Mod &H100
Case 7
rrr = colorh Mod &H100
ggg = (colorh \ &H100) Mod &H100
bbb = (colorh \ &H10000) Mod &H100
Case 8
rrr = colori Mod &H100
ggg = (colori \ &H100) Mod &H100
bbb = (colori \ &H10000) Mod &H100
Case 9
rrr = colorj Mod &H100
ggg = (colorj \ &H100) Mod &H100
bbb = (colorj \ &H10000) Mod &H100

End Select

'Makes RR, GG, and BB negative if they are more than rrr, ggg, or
'     bbb repectively
If RR > rrr Then RR = 0 - RR
If GG > ggg Then GG = 0 - GG
If BB > bbb Then BB = 0 - BB

'Amount to add to RR, GG, and BB for each loop
Dim frst
Dim scnd
Dim thrd
frst = Abs((RR - rrr) / FormGrad.ScaleHeight)
scnd = Abs((GG - ggg) / FormGrad.ScaleHeight)
thrd = Abs((BB - bbb) / FormGrad.ScaleHeight)
'Loops the necessary amount of TimesToFade
Dim diff
For diff = ((CurrentFadeNum - 1) * (FormGrad.ScaleHeight / TimesToFade)) To (FormGrad.ScaleHeight / TimesToFade * CurrentFadeNum)
'Uncomment the next line and put this procedure into form resize
'     for weird effects
'DoEvents
'This is where it selects the effect
Select Case FadeEffect
Case 1
'fades down
FormGrad.Line (0, diff)-(FormGrad.ScaleWidth, diff + 1), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade)), BF
Case 2
'fades right
FormGrad.Line (diff, 0)-(diff + 1, FormGrad.ScaleHeight), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade)), BF
Case 3
'Fades from center
FormGrad.Line (Int(Picture12.ScaleWidth / 2 - diff / 2), Int(Picture13.ScaleHeight / 2 - diff / 2))-(Picture13.ScaleWidth / 2 + diff / 2, Picture13.ScaleHeight / 2 + diff / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade)), B
Case 4
'Fades from the top left to center and back untill it reaches the
'     bottom right
FormGrad.Line (0, diff)-(diff, 0), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth, FormGrad.ScaleHeight - diff)-(FormGrad.ScaleWidth - diff, FormGrad.ScaleHeight), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
Case 5
'Fades from top right to center and back untill it reach the bott
'     om left
FormGrad.Line (0, FormGrad.ScaleHeight - diff)-(diff, FormGrad.ScaleHeight), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth - diff, 0)-(FormGrad.ScaleWidth, diff), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
Case 6
'Fades from top right outwards
FormGrad.Line (0, diff)-(FormGrad.ScaleWidth, diff + 1), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade)), BF
FormGrad.Line (diff, 0)-(diff + 1, FormGrad.ScaleHeight), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade)), BF
Case 7
'Fades from bottom left outwards
FormGrad.Line (0, FormGrad.ScaleHeight - diff)-(FormGrad.ScaleWidth, FormGrad.ScaleHeight - (diff + 1)), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade)), BF
FormGrad.Line (diff, 0)-(diff + 1, FormGrad.ScaleHeight), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade)), BF
Case 8
'Fades from bottom right outwards
FormGrad.Line (0, FormGrad.ScaleHeight - diff)-(FormGrad.ScaleWidth, FormGrad.ScaleHeight - (diff + 1)), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade)), BF
FormGrad.Line (FormGrad.ScaleWidth - diff, 0)-(FormGrad.ScaleWidth - (diff + 1), FormGrad.ScaleHeight), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade)), BF
Case 9
'Fades from top right outwards
FormGrad.Line (0, diff)-(FormGrad.ScaleWidth, diff + 1), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade)), BF
FormGrad.Line (FormGrad.ScaleWidth - diff, 0)-(FormGrad.ScaleWidth - (diff + 1), FormGrad.ScaleHeight), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade)), BF
Case 10
'Fades in a star formation
FormGrad.Line (diff / 2, FormGrad.ScaleHeight - diff / 2)-(FormGrad.ScaleWidth / 2, diff / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line -(FormGrad.ScaleWidth - diff / 2, FormGrad.ScaleHeight - diff / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line -(diff / 2, FormGrad.ScaleHeight / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line -(FormGrad.ScaleWidth - diff / 2, FormGrad.ScaleHeight / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line -(diff / 2, FormGrad.ScaleHeight - diff / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (diff / 2, FormGrad.ScaleHeight / 2)-(FormGrad.ScaleWidth / 3 + diff / 6, FormGrad.ScaleHeight / 8 * 3 + diff / 8), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth - diff / 2, FormGrad.ScaleHeight / 2)-(FormGrad.ScaleWidth / 3 * 2 - diff / 6, FormGrad.ScaleHeight / 8 * 3 + diff / 8), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
Case 11
'Fades in a rainbow fade
'I advice not having the form width be twice as big (or more) as
'     the height
'You could make it into an arc, but i found that makes it take lo
'     nger to load
FormGrad.Circle (FormGrad.ScaleWidth / 2, FormGrad.ScaleHeight), FormGrad.ScaleWidth / 2 - diff / 2, RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
Case 12
'Strange diamond fade
If diff < FormGrad.ScaleWidth / 2 Then
FormGrad.Line (0, diff)-(diff, 0), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth, FormGrad.ScaleHeight - diff)-(FormGrad.ScaleWidth - diff, FormGrad.ScaleHeight), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (0, FormGrad.ScaleHeight - diff)-(diff, FormGrad.ScaleHeight), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth - diff, 0)-(FormGrad.ScaleWidth, diff), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
Else
FormGrad.Line (diff - FormGrad.ScaleWidth / 2, FormGrad.ScaleHeight / 2)-(FormGrad.ScaleWidth / 2, diff - FormGrad.ScaleWidth / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth / 2, diff - FormGrad.ScaleHeight / 2)-(FormGrad.ScaleWidth - (diff - FormGrad.ScaleWidth / 2), FormGrad.ScaleHeight / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (diff - FormGrad.ScaleWidth / 2, FormGrad.ScaleHeight / 2)-(FormGrad.ScaleWidth / 2, FormGrad.ScaleHeight - (diff - FormGrad.ScaleWidth / 2)), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth / 2, FormGrad.ScaleHeight - (diff - FormGrad.ScaleHeight / 2))-(FormGrad.ScaleWidth - (diff - FormGrad.ScaleHeight / 2), FormGrad.ScaleHeight / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
End If
Case 13
'Fades from center in lines, interesting but takes a bit to load
'     on slower computers
FormGrad.Line (FormGrad.ScaleWidth / 2 - diff / 2, FormGrad.ScaleHeight)-(FormGrad.ScaleWidth / 2, FormGrad.ScaleHeight / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth / 2 + diff / 2, FormGrad.ScaleHeight)-(FormGrad.ScaleWidth / 2, FormGrad.ScaleHeight / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (0, FormGrad.ScaleHeight / 2 - diff / 2)-(FormGrad.ScaleWidth / 2, FormGrad.ScaleHeight / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (0, FormGrad.ScaleHeight / 2 + diff / 2)-(FormGrad.ScaleWidth / 2, FormGrad.ScaleHeight / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth / 2 - diff / 2, 0)-(FormGrad.ScaleWidth / 2, FormGrad.ScaleHeight / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth / 2 + diff / 2, 0)-(FormGrad.ScaleWidth / 2, FormGrad.ScaleHeight / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth, FormGrad.ScaleHeight / 2 - diff / 2)-(FormGrad.ScaleWidth / 2, FormGrad.ScaleHeight / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth, FormGrad.ScaleHeight / 2 + diff / 2)-(FormGrad.ScaleWidth / 2, FormGrad.ScaleHeight / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
Case 14
'Strange oval shaped fade
'I recommend lots of colors with this one and setting a backgroun
'     d color
FormGrad.Line (0, diff)-(diff, FormGrad.ScaleHeight), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth, FormGrad.ScaleHeight - diff)-(FormGrad.ScaleWidth - diff, 0), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (0, FormGrad.ScaleHeight - diff)-(diff, 0), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth, diff)-(FormGrad.ScaleWidth - diff, FormGrad.ScaleHeight), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
Case 15
'Makes a cross type fade
FormGrad.Line (-10, -10)-(diff / 2, diff / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade)), B
FormGrad.Line (FormGrad.ScaleWidth + 10, -10)-(FormGrad.ScaleWidth - diff / 2, diff / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade)), B
FormGrad.Line (-10, FormGrad.ScaleHeight + 10)-(diff / 2, FormGrad.ScaleHeight - diff / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade)), B
FormGrad.Line (FormGrad.ScaleWidth + 10, FormGrad.ScaleHeight + 10)-(FormGrad.ScaleWidth - diff / 2, FormGrad.ScaleHeight - diff / 2), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade)), B
Case 16
'Pointy fade
FormGrad.Line (0, diff)-(diff / 2, 0), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth / 2 - diff / 2, FormGrad.ScaleHeight)-(FormGrad.ScaleWidth / 2, FormGrad.ScaleHeight - diff), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth, diff)-(FormGrad.ScaleWidth - diff / 2, 0), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth / 2 + diff / 2, FormGrad.ScaleHeight)-(FormGrad.ScaleWidth / 2, FormGrad.ScaleHeight - diff), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
Case 17
'Fade in the shape of the letter M
FormGrad.Line (0, diff)-(diff / 4, 0), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth / 4 - diff / 4, FormGrad.ScaleHeight)-(FormGrad.ScaleWidth / 4, FormGrad.ScaleHeight - diff), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth / 2, diff)-(FormGrad.ScaleWidth / 2 - diff / 4, 0), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth / 4 + diff / 4, FormGrad.ScaleHeight)-(FormGrad.ScaleWidth / 4, FormGrad.ScaleHeight - diff), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth / 2, diff)-(FormGrad.ScaleWidth / 2 + diff / 4, 0), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth / 4 * 3 - diff / 4, FormGrad.ScaleHeight)-(FormGrad.ScaleWidth / 4 * 3, FormGrad.ScaleHeight - diff), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth, diff)-(FormGrad.ScaleWidth - diff / 4, 0), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
FormGrad.Line (FormGrad.ScaleWidth / 4 * 3 + diff / 4, FormGrad.ScaleHeight)-(FormGrad.ScaleWidth / 4 * 3, FormGrad.ScaleHeight - diff), RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesToFade))
End Select
'If you want to add your own effect, just add the following text
'     for the RGB:
'RGB(Abs(RR + frst * (diff - FormGrad.ScaleHeight / TimesToFade *
'     (CurrentFadeNum - 1)) * TimesToFade), Abs(GG + scnd * (diff - For
'     mGrad.ScaleHeight / TimesToFade * (CurrentFadeNum - 1)) * TimesTo
'     Fade), Abs(BB + thrd * (diff - FormGrad.ScaleHeight / TimesToFade
'     * (CurrentFadeNum - 1)) * TimesToFade))
Next diff
Next CurrentFadeNum
End Function


Public Sub update()
GradEffects Picture1, 2, 2, Picture8.BackColor, Picture9.BackColor
Picture10.Cls
Fad Text3.Text, Text2.Text, Text1.Text, Text6.Text, Text5.Text, Text4.Text, "This is an example of the faded text."
End Sub


Function Fad(R1%, G1%, B1%, R2%, G2%, B2%, TheText$)
' This function fades the text
' If there is error then exit sub
On Error GoTo error
' Calculate the length of the message
textlen$ = Len(TheText)
For i = 1 To textlen$
textdone$ = Left(TheText, i)
lastchr$ = Right(textdone$, 1)
colorX = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
' Change the forecolor of picture box
Picture10.ForeColor = colorX
' Print the character
Picture10.Print lastchr$;
Next i
Exit Function
error:
End Function

VERSION 5.00
Begin VB.Form Frm_Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   10605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11325
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   707
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer_Game 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11880
      Top             =   360
   End
   Begin VB.PictureBox Pic_Game 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'Kein
      Height          =   7200
      Left            =   840
      Picture         =   "Game.frx":0000
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   2
      Top             =   1920
      Width           =   9600
   End
   Begin VB.PictureBox Pic_Ball 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   525
      Left            =   11760
      OLEDropMode     =   2  'Automatisch
      Picture         =   "Game.frx":9604E
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   1
      Top             =   360
      Width           =   525
   End
   Begin VB.PictureBox pic_Balken 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   1500
      Left            =   11520
      Picture         =   "Game.frx":96F56
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   0
      Top             =   360
      Width           =   150
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Ping-Pong"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   9840
      Width           =   990
   End
   Begin VB.Label LabInfo 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "by Günther Brandner 2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9000
      TabIndex        =   6
      Top             =   10200
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "F1 = New Game; F2 = Pause; F3 = Speed; Esc = Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   10200
      Width           =   4020
   End
   Begin VB.Label Lab_Score 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Lab_Lifes 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer, y As Integer, xplus As Boolean, xminus As Boolean, yplus As Boolean, yminus As Boolean, _
y_balken As Integer, Lifes As Integer, LastTime As Long, Score As Long, Speed As Integer, balken_speed, _
FirstLoop As Integer, UserHasPushedF1 As Integer, Loop_Exit As Boolean

Sub Draw_Ball()
    If x >= (Pic_Game.ScaleWidth - Pic_Ball.ScaleWidth) Then
        xminus = True
        If (y + Pic_Ball.ScaleHeight) >= y_balken And y <= y_balken + pic_Balken.ScaleHeight Then
            If FirstLoop = 0 Then
                'Nothing
            Else
                Score = Score + 100
            End If
        Else
            Lifes = Lifes - 1
            LastTime = GetTickCount
            Lab_Lifes.BackColor = vbRed
        End If
    ElseIf x <= 0 Then
        xminus = False
        If (y + Pic_Ball.ScaleHeight) >= y_balken And y <= y_balken + pic_Balken.ScaleHeight Then
            Score = Score + 100
        Else
            Lifes = Lifes - 1
            LastTime = GetTickCount
            Lab_Lifes.BackColor = vbRed
        End If
    End If
    If xminus = True Then
        x = x - Speed
    ElseIf xminus = False Then
        x = x + Speed
    End If
    If y >= Pic_Game.ScaleHeight - Pic_Ball.ScaleHeight Then
        yminus = True
    ElseIf y <= 0 Then
        yminus = False
    End If
    If yminus = True Then
        y = y - Speed
    ElseIf yminus = False Then
        y = y + Speed
    End If
    Lab_Score.Caption = "Score: " & Score
    Lab_Lifes.Caption = Lifes & " Credits"
    If (GetTickCount - LastTime) > 1000 Then _
        Lab_Lifes.BackColor = vbBlack
    BitBlt Pic_Game.hDC, x, y, Pic_Ball.ScaleWidth, Pic_Ball.ScaleHeight, Pic_Ball.hDC, Pic_Ball.ScaleLeft, _
        Pic_Ball.ScaleTop, vbSrcPaint
    FirstLoop = 1
End Sub

Sub Key_Status()
    Dim vkdown As Long
    vkdown = 0
    vkup = 0
    vkf1 = 0
    vkdown = GetKeyState(vbKeyDown)
    If vkdown = -127 Or vkdown = -128 Then
        If y_balken >= Pic_Game.ScaleHeight - pic_Balken.ScaleHeight Then
            'Nothing
        Else
            y_balken = y_balken + balken_speed
        End If
    End If
    vkup = GetKeyState(vbKeyUp)
    If vkup = -127 Or vkup = -128 Then
        If y_balken <= 0 Then
            'Nothing
        Else
            y_balken = y_balken - balken_speed
        End If
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Lifes = 10
        Score = 0
        If Speed = 0 Then Speed = 2
        balken_speed = 20
        Firtsloop = 0
        UserHasPushedF1 = 1
        Loop_Exit = False
        x = 0
        y = 0
    End If
    If KeyCode = vbKeyF2 Then
        If UserHasPushedF1 = 1 And Loop_Exit = False Then
            Loop_Exit = True
        Else
            Loop_Exit = False
    End If
    End If
    If KeyCode = vbKeyEscape Then
        Loop_Exit = True
        reply = MsgBox("Do you really want to exit?", vbQuestion + vbYesNo)
        If reply = vbYes Then
            End
        Else
            Loop_Exit = False
        End If
    End If
    If KeyCode = vbKeyF3 Then
            Loop_Exit = True
            xr = InputBox("Please enter the Speed! - 1-16", "Speed")
            If xr <> "" And xr >= 1 And xr <= 16 Then
                Speed = xr
                If Speed >= 8 Then
                    balken_speed = 30
                Else
                    balken_speed = 20
                End If
            End If
        If UserHasPushedF1 Then Loop_Exit = False

    End If
    Call Game_Loop
End Sub

Private Sub Form_Load()
    FirstLoop = 0
    UserHasPushedF1 = 0
End Sub
    
Sub Game_Loop()
    Do While Loop_Exit = False
            Pic_Game.Cls
            Call Key_Status
            Call Draw_Ball
            Call Draw_Balken
            Pic_Game.Refresh
            If Lifes = 0 Then
                MsgBox "You're the looser!", vbInformation
                Pic_Game.Cls
                Loop_Exit = True
            End If
            DoEvents
    Loop
End Sub

Sub Draw_Balken()
    BitBlt Pic_Game.hDC, 0, y_balken, pic_Balken.ScaleWidth, pic_Balken.ScaleHeight, pic_Balken.hDC, pic_Balken.ScaleLeft, _
        pic_Balken.ScaleTop, vbSrcCopy
    BitBlt Pic_Game.hDC, Pic_Game.ScaleWidth - pic_Balken.ScaleWidth, y_balken, pic_Balken.ScaleWidth, pic_Balken.ScaleHeight, pic_Balken.hDC, pic_Balken.ScaleLeft, _
        pic_Balken.ScaleTop, vbSrcCopy
End Sub

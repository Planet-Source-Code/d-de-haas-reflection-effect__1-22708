VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Reflection Demo"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   ScaleHeight     =   8250
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   3375
      Left            =   6360
      ScaleHeight     =   3315
      ScaleWidth      =   2955
      TabIndex        =   7
      Top             =   240
      Width           =   3015
      Begin VB.CommandButton Command2 
         Caption         =   "&End"
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   2760
         Width           =   855
      End
      Begin VB.HScrollBar HScrollSize 
         Height          =   255
         Left            =   240
         Max             =   50
         Min             =   1
         TabIndex        =   1
         Top             =   1080
         Value           =   20
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Pause"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2760
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Small Reflection"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Large Reflection"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   1575
      End
      Begin VB.HScrollBar HScrollSpeed 
         Height          =   255
         Left            =   240
         Max             =   400
         TabIndex        =   0
         Top             =   480
         Value           =   50
         Width           =   2055
      End
      Begin VB.HScrollBar HScrollAngle 
         Height          =   255
         Left            =   240
         Max             =   200
         TabIndex        =   2
         Top             =   1680
         Value           =   100
         Width           =   2055
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "?"
         Height          =   195
         Left            =   2400
         TabIndex        =   13
         Top             =   1680
         Width           =   90
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "?"
         Height          =   195
         Left            =   2400
         TabIndex        =   12
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "?"
         Height          =   195
         Left            =   2400
         TabIndex        =   11
         Top             =   480
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Wave Size"
         Height          =   195
         Left            =   840
         TabIndex        =   10
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Wave Speed"
         Height          =   195
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Wave Angle"
         Height          =   195
         Left            =   840
         TabIndex        =   8
         Top             =   1440
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3810
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3810
      ScaleWidth      =   5850
      TabIndex        =   6
      Top             =   240
      Width           =   5850
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare variables for the program
Private FlipPic As tNEWBMP

Public bGo As Boolean
Public bEndProg As Boolean
Public bValueChanged As Boolean
    
Private Sub Command1_Click()
    ' Pause or not to pause
    bGo = Not bGo
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
' For the best look, set the properties of the picture
' Appearance = 0  ' Flat
' BorderStyle = 0 ' None
    
    Me.ScaleMode = vbPixels
    Me.Visible = True
        
    With Picture1
        .ScaleMode = vbPixels
        .Visible = True
        
        ' Update the controlpanel
        Label4.Caption = HScrollSpeed.Value
        Label5.Caption = HScrollSize.Value
        Label6.Caption = HScrollAngle.Value
        
        Me.Refresh
        
        ' Flip the picture and copy it into memory
        FlipPic = FlipBitmap(.hdc, .ScaleWidth, .ScaleHeight)
        
        If FlipPic.hdc <> 0 Then ' Check if the picture was build correctly
            
            ' start waving
            StartWave .Left, .Top + .Height
            
        End If
    End With
End Sub

Public Sub StartWave(ByVal X As Long, ByVal Y As Long)
    Dim lrtnV As Long
    
    Dim Yloop As Long
    Dim Yloop1 As Long
    
    Dim Xs As Single
    Dim Xn As Long

    Dim PI As Single
    Dim Degrees As Single
    Dim WaveSpeed As Single
    Dim Amplitude As Single
    Dim AmplitudeMax As Single
    Dim AmpStep As Single
    
    Dim iLargePic As Integer
    
    ' Calculate PI
    PI = Atn(1) * 4
            
    ' Set for start
    WaveSpeed = 0
    bGo = True
    bValueChanged = True
    
    With FlipPic
        Do
            If bValueChanged Then
                ' Read values from the controlpanel
                
                If Option1(0).Value Then
                    iLargePic = 1
                Else
                    iLargePic = 2
                End If
                ' Max wave size
                AmplitudeMax = HScrollSize.Value
                
                If FlipPic.lHeight <> 0 Then ' Just in case
                
                    ' Calculate steps for a nice smooth wave
                    If iLargePic = 2 Then ' Switch between a large of small pic
                        AmpStep = AmplitudeMax / FlipPic.lHeight
                    Else
                        AmpStep = AmplitudeMax / (FlipPic.lHeight / 2)
                    End If
                                                        
                End If
                
                Me.Cls
                
                bValueChanged = False
            End If
        
            ' Set for start at each wave
            Amplitude = 0
            Xs = X
            Yloop1 = 0
            
            ' Loop though the flipped picture line by line
            For Yloop = 0 To .lHeight Step iLargePic
                Degrees = (Yloop + WaveSpeed) * (PI / 180) ' convert angle from radians to degrees
                
                Xn = Xs + (Amplitude * Sin(5 * Degrees)) ' calculate wave
            
                ' Draw line
                lrtnV = BitBlt(Me.hdc, Xn, Y + Yloop1, .lWidth, 1, .hdc, 0, Yloop, SRCCOPY)
                If lrtnV = 0 Then
                    MsgBox "The bitmap is not build correctly.", vbCritical, "Error (BitBlt)"
                
                    Exit Sub
                End If
                
                ' Bend the reflexion to the right
                Xs = Xs + (HScrollAngle.Value / 100) ' Read the value from the controlpanel
                
                ' Waving
                Amplitude = Amplitude + AmpStep
                Yloop1 = Yloop1 + 1
            Next Yloop
            
            ' Speedup the wave
            WaveSpeed = WaveSpeed + (HScrollSpeed.Value / 100) ' Read the value from the controlpanel
            
            Me.Refresh
            
            Do
                ' Pauze loop
                DoEvents
            Loop Until bGo Or bEndProg
        
        Loop Until bEndProg ' Loop until program stops
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bEndProg = True
    
    ' Clean up
    FreeResource FlipPic
    End
    
End Sub

Private Sub HScrollAngle_Change()
    bValueChanged = True
    
    Label6.Caption = HScrollAngle.Value
End Sub

Private Sub HScrollSize_Change()
    bValueChanged = True
    
    Label5.Caption = HScrollSize.Value
End Sub

Private Sub HScrollSpeed_Change()
    Label4.Caption = HScrollSpeed.Value
End Sub

Private Sub Option1_Click(Index As Integer)
    bValueChanged = True
End Sub

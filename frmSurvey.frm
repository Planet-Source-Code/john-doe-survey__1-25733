VERSION 5.00
Begin VB.Form frmSurvey 
   Caption         =   "Survey"
   ClientHeight    =   6390
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "frmSurvey.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8880
   Begin VB.Frame Frame13 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   840
      TabIndex        =   89
      Top             =   0
      Width           =   7935
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "This is the title of your survey."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   91
         Top             =   240
         Width           =   4170
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "How do you feel about the subjects below as they pertain to your business? Please mark the appropriate answer."
         Height          =   555
         Left            =   840
         TabIndex        =   90
         Top             =   720
         Width           =   5655
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   -120
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   8400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame7 
      Caption         =   "Subject #7"
      Height          =   735
      Left            =   840
      TabIndex        =   76
      Top             =   6600
      Width           =   7455
      Begin VB.OptionButton Option7 
         Caption         =   "Strongly agree"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Agree"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Somewhat agree"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Disagree"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Does not apply"
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Subject #8"
      Height          =   735
      Left            =   840
      TabIndex        =   70
      Top             =   7440
      Width           =   7455
      Begin VB.OptionButton Option8 
         Caption         =   "Strongly agree"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Agree"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Somewhat agree"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Disagree"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Does not apply"
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Subject #9"
      Height          =   735
      Left            =   840
      TabIndex        =   64
      Top             =   8280
      Width           =   7455
      Begin VB.OptionButton Option9 
         Caption         =   "Does not apply"
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Disagree"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Somewhat agree"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Agree"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Strongly agree"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Subject #10"
      Height          =   735
      Left            =   840
      TabIndex        =   58
      Top             =   9120
      Width           =   7455
      Begin VB.OptionButton Option10 
         Caption         =   "Strongly agree"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Agree"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Somewhat agree"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Disagree"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Does not apply"
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Subject #6"
      Height          =   735
      Left            =   840
      TabIndex        =   52
      Top             =   5760
      Width           =   7455
      Begin VB.OptionButton Option6 
         Caption         =   "Does not apply"
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Disagree"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Somewhat agree"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Agree"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Strongly agree"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Subject #1"
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   7455
      Begin VB.OptionButton Option1 
         Caption         =   "Strongly agree"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Agree"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Somewhat agree"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Disagree"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Does not apply"
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Subject #2"
      Height          =   735
      Left            =   840
      TabIndex        =   41
      Top             =   2400
      Width           =   7455
      Begin VB.OptionButton Option2 
         Caption         =   "Strongly agree"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Agree"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Somewhat agree"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Disagree"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Does not apply"
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Subject #3"
      Height          =   735
      Left            =   840
      TabIndex        =   35
      Top             =   3240
      Width           =   7455
      Begin VB.OptionButton Option3 
         Caption         =   "Does not apply"
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Disagree"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Somewhat agree"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Agree"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Strongly agree"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Subject #4"
      Height          =   735
      Left            =   840
      TabIndex        =   29
      Top             =   4080
      Width           =   7455
      Begin VB.OptionButton Option4 
         Caption         =   "Strongly agree"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Agree"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Somewhat agree"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Disagree"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Does not apply"
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Subject #5"
      Height          =   735
      Left            =   840
      TabIndex        =   23
      Top             =   4920
      Width           =   7455
      Begin VB.OptionButton Option5 
         Caption         =   "Does not apply"
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Disagree"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Somewhat agree"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Agree"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Strongly agree"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1455
      Left            =   9240
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1440
      Width           =   255
   End
   Begin VB.Frame frmUserInfo 
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   240
      TabIndex        =   1
      Top             =   10080
      Width           =   8055
      Begin VB.TextBox txtAddress3 
         Height          =   375
         Left            =   3480
         TabIndex        =   10
         Top             =   3720
         Width           =   4815
      End
      Begin VB.TextBox txtAddress2 
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Top             =   3360
         Width           =   4815
      End
      Begin VB.TextBox txtAddress1 
         Height          =   375
         Left            =   3480
         TabIndex        =   8
         Top             =   3000
         Width           =   4815
      End
      Begin VB.OptionButton optServices 
         Caption         =   "No"
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   5280
         Width           =   615
      End
      Begin VB.OptionButton optServices 
         Caption         =   "Yes"
         Height          =   315
         Index           =   0
         Left            =   4080
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   5280
         Width           =   675
      End
      Begin VB.TextBox txtEmail 
         Height          =   405
         Left            =   3480
         TabIndex        =   12
         Top             =   4680
         Width           =   4815
      End
      Begin VB.TextBox txtPhone 
         Height          =   405
         Left            =   3480
         TabIndex        =   11
         Top             =   4200
         Width           =   4815
      End
      Begin VB.TextBox txtCompany 
         Height          =   405
         Left            =   3480
         TabIndex        =   7
         Top             =   2520
         Width           =   4815
      End
      Begin VB.TextBox txtName 
         Height          =   405
         Left            =   3480
         TabIndex        =   6
         Top             =   2040
         Width           =   4815
      End
      Begin VB.TextBox txtLocation 
         Height          =   405
         Left            =   3480
         TabIndex        =   5
         Top             =   1560
         Width           =   4815
      End
      Begin VB.TextBox txtType 
         Height          =   372
         Left            =   3480
         TabIndex        =   4
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox txtConcern2 
         Height          =   372
         Left            =   3480
         TabIndex        =   3
         Top             =   480
         Width           =   4815
      End
      Begin VB.TextBox txtConcern1 
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   0
         Width           =   4815
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "Submit"
         Height          =   375
         Left            =   3480
         TabIndex        =   83
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label lblAddress3 
         Alignment       =   1  'Right Justify
         Caption         =   "Address (line 3):"
         Height          =   255
         Left            =   2040
         TabIndex        =   88
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label lblAddress2 
         Alignment       =   1  'Right Justify
         Caption         =   "Address (line 2):"
         Height          =   255
         Left            =   2040
         TabIndex        =   87
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label lblServices 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Are you currently using our services?"
         Height          =   195
         Left            =   1380
         TabIndex        =   86
         Top             =   5280
         Width           =   2595
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         Caption         =   "Email:"
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label lblPhone 
         Alignment       =   1  'Right Justify
         Caption         =   "Telephone:"
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label lblAddress1 
         Alignment       =   1  'Right Justify
         Caption         =   "Address (line 1):"
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         Caption         =   "Company:"
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Name/Title:"
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Caption         =   "Location:"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "Type of business:"
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblConcerns 
         Alignment       =   1  'Right Justify
         Caption         =   "List two of your top concerns not covered in this survey:"
         Height          =   495
         Left            =   720
         TabIndex        =   13
         Top             =   240
         Width           =   2655
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmSurvey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim subject1 As Integer
Dim subject2 As Integer
Dim subject3 As Integer
Dim subject4 As Integer
Dim subject5 As Integer
Dim subject6 As Integer
Dim subject7 As Integer
Dim subject8 As Integer
Dim subject9 As Integer
Dim subject10 As Integer
Dim services As String

Dim iFullFormHeigth As Integer
Dim iFullFormWidth As Integer
Dim oldvPos As Integer
Dim oldhPos As Integer
Private Sub Form_Load()
GetFullSize
Form_Resize
End Sub

Private Sub GetFullSize()
Dim ctl As Control
Dim fullhtemp As Integer
Dim fullvtemp As Integer

fullhtemp = 0
fullvtemp = 0
For Each ctl In Me.Controls
        If ctl.Top + ctl.Height > fullvtemp Then fullvtemp = ctl.Top + ctl.Height
        If ctl.Left + ctl.Width > fullhtemp Then fullhtemp = ctl.Left + ctl.Width
Next
iFullFormHeigth = fullvtemp + HScroll1.Height
iFullFormWidth = fullhtemp + VScroll1.Width
End Sub

Private Sub Form_Resize()

VScroll1.Left = Me.Width - (1.45 * VScroll1.Width)
HScroll1.Top = Me.Height - (2.45 * HScroll1.Height)

Picture1.Left = VScroll1.Left
Picture1.Top = HScroll1.Top

VScroll1.Enabled = (iFullFormHeigth - Me.Height) >= 0

If Me.ScaleHeight > HScroll1.Height And Me.Width > VScroll1.Width Then
    
    If VScroll1.Enabled Then
        With VScroll1
            .Height = Me.ScaleHeight - HScroll1.Height
            .Min = 0
            .Max = iFullFormHeigth - Me.Height
            .SmallChange = Screen.TwipsPerPixelY * 10
            .LargeChange = Me.ScaleHeight - HScroll1.Height
        End With

    Else: VScroll1.Height = Me.ScaleHeight - HScroll1.Height
    End If

    HScroll1.Enabled = (iFullFormWidth - Me.Width) >= 0
    If HScroll1.Enabled Then
        With HScroll1
            .Width = Me.ScaleWidth - VScroll1.Width
            .Min = 0
            .Max = iFullFormWidth - Me.Width
            .SmallChange = Screen.TwipsPerPixelX * 10
            .LargeChange = Me.ScaleWidth - VScroll1.Width
        End With

    Else: HScroll1.Width = Me.ScaleWidth - VScroll1.Width
    End If
End If
End Sub

Private Sub pScrollForm()
Dim ctl As Control

For Each ctl In Me.Controls
    If Not (TypeOf ctl Is VScrollBar) And _
        Not (TypeOf ctl Is PictureBox) And _
               Not (TypeOf ctl Is OptionButton) And _
               Not (TypeOf ctl Is TextBox) And _
                Not (TypeOf ctl Is Label) And _
                Not (TypeOf ctl Is CommandButton) And _
        Not (TypeOf ctl Is HScrollBar) Then
        ctl.Top = ctl.Top + oldvPos - VScroll1.Value
        ctl.Left = ctl.Left + oldhPos - HScroll1.Value
    End If
Next

oldvPos = VScroll1.Value
oldhPos = HScroll1.Value
End Sub

Private Sub VScroll1_Change()
    Call pScrollForm
End Sub

Private Sub VScroll1_Scroll()
    Call pScrollForm
End Sub

Private Sub HScroll1_Change()
    Call pScrollForm
End Sub

Private Sub HScroll1_Scroll()
    Call pScrollForm
End Sub

Private Sub Option1_Click(Index As Integer)
    Frame1.Tag = Index
 Select Case Index
    Case 0
    subject1 = "1"
    Case 1
    subject1 = "2"
    Case 2
    subject1 = "3"
    Case 3
    subject1 = "4"
    Case 4
    subject1 = "5"
    End Select
End Sub

Private Sub Option2_Click(Index As Integer)
        Frame2.Tag = Index
    Select Case Index
    Case 0
    subject2 = "1"
    Case 1
    subject2 = "2"
    Case 2
    subject2 = "3"
    Case 3
    subject2 = "4"
    Case 4
    subject2 = "5"
    End Select
End Sub

Private Sub Option3_Click(Index As Integer)
    Frame3.Tag = Index
    Select Case Index
    Case 0
    subject3 = "1"
    Case 1
    subject3 = "2"
    Case 2
    subject3 = "3"
    Case 3
    subject3 = "4"
    Case 4
    subject3 = "5"
    End Select
End Sub

Private Sub Option4_Click(Index As Integer)
    Frame4.Tag = Index
    Select Case Index
    Case 0
    subject4 = "1"
    Case 1
    subject4 = "2"
    Case 2
    subject4 = "3"
    Case 3
    subject4 = "4"
    Case 4
    subject4 = "5"
    End Select
End Sub

Private Sub Option5_Click(Index As Integer)
    Frame5.Tag = Index
    Select Case Index
    Case 0
    subject5 = "1"
    Case 1
    subject5 = "2"
    Case 2
    subject5 = "3"
    Case 3
    subject5 = "4"
    Case 4
    subject5 = "5"
    End Select
End Sub
Private Sub Option6_Click(Index As Integer)
    Frame6.Tag = Index
    Select Case Index
    Case 0
    subject6 = "1"
    Case 1
    subject6 = "2"
    Case 2
    subject6 = "3"
    Case 3
    subject6 = "4"
    Case 4
    subject6 = "5"
    End Select
End Sub
Private Sub Option7_Click(Index As Integer)
    Frame7.Tag = Index
    Select Case Index
    Case 0
    subject7 = "1"
    Case 1
    subject7 = "2"
    Case 2
    subject7 = "3"
    Case 3
    subject7 = "4"
    Case 4
    subject7 = "5"
    End Select
End Sub
Private Sub Option8_Click(Index As Integer)
    Frame8.Tag = Index
    Select Case Index
    Case 0
    subject8 = "1"
    Case 1
    subject8 = "2"
    Case 2
    subject8 = "3"
    Case 3
    subject8 = "4"
    Case 4
    subject8 = "5"
    End Select
End Sub
Private Sub Option9_Click(Index As Integer)
    Frame9.Tag = Index
    Select Case Index
    Case 0
    subject9 = "1"
    Case 1
    subject9 = "2"
    Case 2
    subject9 = "3"
    Case 3
    subject9 = "4"
    Case 4
    subject9 = "5"
    End Select
End Sub
Private Sub Option10_Click(Index As Integer)
    Frame10.Tag = Index
    Select Case Index
    Case 0
    subject10 = "1"
    Case 1
    subject10 = "2"
    Case 2
    subject10 = "3"
    Case 3
    subject10 = "4"
    Case 4
    subject10 = "5"
    End Select
End Sub
Private Sub optServices_Click(Index As Integer)
    Select Case Index
    Case 0
    services = "Yes"
    Case 1
    services = "No"
    End Select
End Sub

Private Sub cmdSubmit_Click()
Dim MyFile

   If Frame1.Tag = "" Then
    MsgBox "Please complete Subject #1.", vbExclamation, "Survey"
       
   ElseIf Frame2.Tag = "" Then
    MsgBox "Please complete Subject #2.", vbExclamation, "Survey"
       
   ElseIf Frame3.Tag = "" Then
    MsgBox "Please complete Subject #3.", vbExclamation, "Survey"
       
   ElseIf Frame4.Tag = "" Then
    MsgBox "Please complete Subject #4.", vbExclamation, "Survey"
       
   ElseIf Frame5.Tag = "" Then
    MsgBox "Please complete Subject #5.", vbExclamation, "Survey"
       
   ElseIf Frame6.Tag = "" Then
    MsgBox "Please complete Subject #6.", vbExclamation, "Survey"
       
   ElseIf Frame7.Tag = "" Then
    MsgBox "Please complete Subject #7.", vbExclamation, "Survey"
       
   ElseIf Frame8.Tag = "" Then
    MsgBox "Please complete Subject #8.", vbExclamation, "Survey"
    
   ElseIf Frame9.Tag = "" Then
    MsgBox "Please complete Subject #9.", vbExclamation, "Survey"

   ElseIf Frame10.Tag = "" Then
    MsgBox "Please complete Subject #10.", vbExclamation, "Survey"
     
   ElseIf txtType.Text = "" Then
    MsgBox "Please fill in the business type.", vbExclamation, "Survey"
      txtType.SetFocus

    ElseIf txtLocation.Text = "" Then
    MsgBox "Please fill in the location.", vbExclamation, "Survey"
       txtLocation.SetFocus

      Else

    MyFile = "C:\Survey.txt"
    Open MyFile For Append As #1
    Print #1, subject1 & Chr(44) & subject2 & Chr(44) & subject3 & Chr(44) & subject4 & Chr(44) & subject5 & Chr(44) & subject6 & Chr(44) & subject7 & Chr(44) & subject8 & Chr(44) & subject9 & Chr(44) & subject10 & Chr(44) & _
    vbCrLf & "Concern 1: " & txtConcern1.Text & _
    vbCrLf & "Concern 2: " & txtConcern2.Text & _
    vbCrLf & "Type of Business: " & txtType.Text & _
    vbCrLf & "Location: " & txtLocation.Text & _
    vbCrLf & "Contact Info (next line):" & _
    vbCrLf & txtName.Text & Chr(44) & txtCompany.Text & Chr(44) & txtAddress1.Text & Chr(44); txtAddress2.Text & Chr(44) & txtAddress3.Text & Chr(44) & txtPhone.Text & Chr(44) & txtEmail.Text & _
    vbCrLf & "Do you use our services? : " & services
    Close #1
    
    MsgBox "Your responses have been submitted." & vbCrLf & "Thank you for taking our survey!", vbExclamation, "Survey"

   Call ClearForm(Me)
   
   End If
End Sub

Public Sub ClearForm(frmSurvey As Form)
Dim MyControl As Control
  
  For Each MyControl In frmSurvey.Controls
      If TypeOf MyControl Is TextBox Then MyControl.Text = ""
      If TypeOf MyControl Is OptionButton Then MyControl.Value = "False"
      If TypeOf MyControl Is Frame Then MyControl.Tag = ""
   Next
End Sub

 Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Do you want to close the survey?", vbQuestion + vbYesNo, "Close Survey") = vbYes Then
        Unload Me
    Else
        Cancel = True
    End If
End Sub

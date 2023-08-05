VERSION 5.00
Begin VB.Form Exercise1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WordCount"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8880
   FillColor       =   &H00FFC0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "WordCount.frx":0000
   ScaleHeight     =   4.667
   ScaleMode       =   5  'Inch
   ScaleWidth      =   6.167
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Count"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "<<<Go Back"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   6240
      Width           =   2055
   End
End
Attribute VB_Name = "Exercise1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Dim tempArray()     As String
        Dim lngWordCount    As Long
        Dim lngCharCount    As Long
        Dim lngCharCountS    As Long
        
        tempArray = Split(Trim$(Text1.Text), " ")
        
        lngWordCount = UBound(tempArray) + 1
        lngCharCountS = Len(Replace(Text1.Text, " ", ""))

        MsgBox "Word Count = " & lngWordCount & vbNewLine & _
                "Character count = " & lngCharCountS
End Sub

Private Sub Label1_Click()
Exercise1.Hide
MainForm.Show
End Sub

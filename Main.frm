VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student Loans Calculator"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton github 
      Caption         =   "Github Link"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   14
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Loan_Totoal_Year_Text 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Text            =   "10"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Interest_Rate_Text 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Text            =   "4.9"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Interest 
      Height          =   2175
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   4320
      Width           =   3615
   End
   Begin VB.TextBox Last_Year_Principal 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Text            =   "1000"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Common_Year_Principal 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Text            =   "1000"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox No_Principal_Year 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Text            =   "3"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Total_Loans_text 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Text            =   "7000"
      ToolTipText     =   "输入贷款总额"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Most_Interest 
      Caption         =   "按规定还款利息计算"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Loan_Totoal_Year 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "贷款年限"
      Height          =   180
      Left            =   600
      TabIndex        =   13
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Interest_Rate 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "利率"
      Height          =   180
      Left            =   840
      TabIndex        =   12
      Top             =   720
      Width           =   360
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "最后一年还本金"
      Height          =   180
      Left            =   360
      TabIndex        =   10
      Top             =   2880
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "中间每年还本金金额"
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "不用还本金年数"
      Height          =   180
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label Total_Loans 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "贷款总额"
      Height          =   180
      Left            =   720
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '变量必须定义才能使用
'ShellExecute函数需要添加下面的语句
''''这一段是ShellExecute函数需要加入的一些操作
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'''''ShellExecute函数用来打开文件、程序、链接等骚操作
Dim return_num As Integer  '用于ShellExecute，函数返回值，没有返回值接收会报错
Dim loan_totoal_year_ As Integer
'测试打开http://note.youdao.com/noteshare?id=3bfdcedbb35f23c29d5fbc1a4f0c8a58
'return_num = ShellExecute(Me.hwnd, "open", "http://note.youdao.com/noteshare?id=3bfdcedbb35f23c29d5fbc1a4f0c8a58 ", "", "", 1)
Dim total_interest As Double  '利息和统计


Private Sub Most_Interest_Click()
loan_totoal_year_ = Val(Loan_Totoal_Year_Text.Text)
Dim i, j As Integer
j = 0
For i = 1 To loan_totoal_year_
    Dim interest_temp As Double
    If i = 1 Then
        Interest.Text = "第" & i & "年利息=" & VBA.Format(Val(Total_Loans_text.Text) * _
                        Val(Interest_Rate_Text.Text) / 100 * 110 / 365, "#0.0") & "元    "
        interest_temp = VBA.Format(Val(Total_Loans_text.Text) * Val(Interest_Rate_Text.Text) _
                        / 100 * 110 / 365, "#0.0")
    ElseIf (i > 1 And i <= Val(No_Principal_Year.Text)) Then
        Interest.Text = Interest.Text & "第" & i & "年利息=" & VBA.Format(Val(Total_Loans_text.Text) _
                        * Val(Interest_Rate_Text.Text) / 100, "#0.0") & "元    "
        interest_temp = VBA.Format(Val(Total_Loans_text.Text) * Val(Interest_Rate_Text.Text) / 100, "#0.0")
    ElseIf (i > Val(No_Principal_Year.Text) And i <= Val(Loan_Totoal_Year_Text.Text) - 1) Then
        Interest.Text = Interest.Text & "第" & i & "年利息=" & VBA.Format((Val(Total_Loans_text.Text) _
                        - j * Val(Common_Year_Principal.Text)) * Val(Interest_Rate_Text.Text) / 100, "#0.0") _
                        & "元    "
        interest_temp = VBA.Format((Val(Total_Loans_text.Text) - j * Val(Common_Year_Principal.Text)) _
                        * Val(Interest_Rate_Text.Text) / 100, "#0.0")
        j = j + 1
    Else
        Interest.Text = Interest.Text & "第" & i & "年利息=" & VBA.Format((Val(Total_Loans_text.Text) _
                        - j * Val(Common_Year_Principal.Text)) * Val(Interest_Rate_Text.Text) / 100, "#0.0") _
                        & "元    "
        interest_temp = VBA.Format((Val(Total_Loans_text.Text) - j * Val(Common_Year_Principal.Text)) _
                        * Val(Interest_Rate_Text.Text) / 100, "#0.0")
    End If
    
    total_interest = total_interest + interest_temp
    
    If (i Mod 2 = 0) And i <> loan_totoal_year_ Then
            Interest.Text = Interest.Text & vbCrLf
    End If
Next i
Interest.Text = Interest.Text & vbCrLf & "贷款总利息=" & total_interest & "元"
total_interest = 0
End Sub

Private Sub github_Click()
return_num = ShellExecute(Me.hwnd, "open", "https://github.com/geekerboy/Student_Loans_Calculator.git", _
"", "", 1)
End Sub





VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student Loans Calculator"
   ClientHeight    =   6705
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3720
      Top             =   240
   End
   Begin VB.TextBox Result_Repay_Advance 
      Height          =   1455
      Left            =   4800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   44
      Top             =   4320
      Width           =   3975
   End
   Begin VB.CommandButton Repay_Advance 
      Caption         =   "计算提前还款利息"
      Height          =   615
      Left            =   5760
      TabIndex        =   43
      Top             =   3480
      Width           =   2175
   End
   Begin VB.ComboBox Today_Request 
      Height          =   300
      Left            =   6480
      TabIndex        =   31
      Text            =   "是"
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox Predict_Day 
      Height          =   300
      Left            =   7080
      TabIndex        =   30
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox Predict_Month 
      Height          =   300
      Left            =   5880
      TabIndex        =   29
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Reapy_Principal_Text 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6480
      TabIndex        =   26
      Text            =   "1000"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Rate_Text 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6480
      TabIndex        =   21
      Text            =   "4.9"
      Top             =   1080
      Width           =   855
   End
   Begin VB.ComboBox Month_Choose 
      Height          =   300
      Left            =   6480
      TabIndex        =   20
      Text            =   "12"
      Top             =   1560
      Width           =   855
   End
   Begin VB.ComboBox Loan_Year 
      Height          =   300
      Left            =   5880
      TabIndex        =   19
      Text            =   "既非第一年也非最后一年"
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox Remaining_Principle_Text 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6480
      TabIndex        =   18
      Text            =   "6000"
      Top             =   120
      Width           =   855
   End
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
      Left            =   6720
      TabIndex        =   14
      Top             =   6120
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
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   4320
      Width           =   3855
   End
   Begin VB.TextBox Last_Year_Principal 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Text            =   "1000"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Common_Year_Principal 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Text            =   "1000"
      Top             =   2040
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
   Begin VB.Label Time 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "time"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   46
      Top             =   3840
      Width           =   600
   End
   Begin VB.Label Date 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "date"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   45
      Top             =   3480
      Width           =   600
   End
   Begin VB.Label month 
      AutoSize        =   -1  'True
      Caption         =   "月"
      Height          =   180
      Index           =   1
      Left            =   7440
      TabIndex        =   42
      Top             =   1680
      Width           =   180
   End
   Begin VB.Label yuan 
      AutoSize        =   -1  'True
      Caption         =   "元"
      Height          =   180
      Index           =   4
      Left            =   7440
      TabIndex        =   41
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label year 
      AutoSize        =   -1  'True
      Caption         =   "年"
      Height          =   180
      Index           =   1
      Left            =   3120
      TabIndex        =   40
      Top             =   1680
      Width           =   180
   End
   Begin VB.Label year 
      AutoSize        =   -1  'True
      Caption         =   "年"
      Height          =   180
      Index           =   0
      Left            =   3120
      TabIndex        =   39
      Top             =   1200
      Width           =   180
   End
   Begin VB.Label rate 
      AutoSize        =   -1  'True
      Caption         =   "%"
      Height          =   180
      Index           =   1
      Left            =   7440
      TabIndex        =   38
      Top             =   1200
      Width           =   90
   End
   Begin VB.Label rate 
      AutoSize        =   -1  'True
      Caption         =   "%"
      Height          =   180
      Index           =   0
      Left            =   3120
      TabIndex        =   37
      Top             =   720
      Width           =   90
   End
   Begin VB.Label yuan 
      AutoSize        =   -1  'True
      Caption         =   "元"
      Height          =   180
      Index           =   3
      Left            =   7440
      TabIndex        =   36
      Top             =   240
      Width           =   180
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   90
   End
   Begin VB.Label yuan 
      AutoSize        =   -1  'True
      Caption         =   "元"
      Height          =   180
      Index           =   2
      Left            =   3120
      TabIndex        =   34
      Top             =   2640
      Width           =   180
   End
   Begin VB.Label yuan 
      AutoSize        =   -1  'True
      Caption         =   "元"
      Height          =   180
      Index           =   1
      Left            =   3120
      TabIndex        =   33
      Top             =   2160
      Width           =   180
   End
   Begin VB.Label yuan 
      AutoSize        =   -1  'True
      Caption         =   "元"
      Height          =   180
      Index           =   0
      Left            =   3120
      TabIndex        =   32
      Top             =   240
      Width           =   180
   End
   Begin VB.Label day 
      AutoSize        =   -1  'True
      Caption         =   "日"
      Height          =   180
      Left            =   7920
      TabIndex        =   28
      Top             =   3000
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label month 
      AutoSize        =   -1  'True
      Caption         =   "月"
      Height          =   180
      Index           =   0
      Left            =   6720
      TabIndex        =   27
      Top             =   3000
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Predict_date_ 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "预计申请归还日期"
      Height          =   180
      Left            =   4200
      TabIndex        =   25
      Top             =   3000
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Today_Request_ 
      AutoSize        =   -1  'True
      Caption         =   "是否今天申请还款"
      Height          =   180
      Left            =   4200
      TabIndex        =   24
      Top             =   2520
      Width           =   1440
   End
   Begin VB.Label Reapy_Principal_ 
      AutoSize        =   -1  'True
      Caption         =   "打算归还本金"
      Height          =   180
      Left            =   4320
      TabIndex        =   23
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Label Rate_ 
      AutoSize        =   -1  'True
      Caption         =   "利率"
      Height          =   180
      Left            =   4800
      TabIndex        =   22
      Top             =   1200
      Width           =   360
   End
   Begin VB.Label Last_Deduction 
      AutoSize        =   -1  'True
      Caption         =   "上一次扣款月份"
      Height          =   180
      Left            =   4320
      TabIndex        =   17
      Top             =   1560
      Width           =   1260
   End
   Begin VB.Label Loan_Yeear_ 
      AutoSize        =   -1  'True
      Caption         =   "还款年度"
      Height          =   180
      Left            =   4680
      TabIndex        =   16
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Remaining_Principle_ 
      AutoSize        =   -1  'True
      Caption         =   "当前剩余本金"
      Height          =   180
      Left            =   4560
      TabIndex        =   15
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label Loan_Totoal_Year 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "贷款年限"
      Height          =   180
      Left            =   600
      TabIndex        =   13
      Top             =   1200
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
      Left            =   240
      TabIndex        =   10
      Top             =   2640
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
      Top             =   1680
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
   Begin VB.Menu Deep_Calculate 
      Caption         =   "深度计算"
      Index           =   0
      Begin VB.Menu Fixed_Principal 
         Caption         =   "每年还定额（待开发）"
         Index           =   1
      End
   End
   Begin VB.Menu Calculate_Info 
      Caption         =   "计算说明"
      Index           =   0
      Begin VB.Menu Settlement_Cycle 
         Caption         =   "结算周期说明"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'log
'用菜单添加利息计算的相关说明
'添加单位，调整一下排版细节
'添加时间显示

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
Dim month_now As Integer '获取系统当前月份
Dim day_now As Integer '获取系统当前具体多少日
Dim month_reall As Integer  '实际参与利息天数计算的月份变量
Dim day_reall As Integer  '实际参与利息天数计算的日期变量
Dim msg_return As Integer 'msg消息返回值


Private Sub Loan_Year_Click()
Select Case Loan_Year.Text
Case "合同第一年度还款"
    '''实际天数为9.1到12.20号
    Month_Choose.Clear
    Month_Choose.AddItem (9)
    Month_Choose.AddItem (10)
    return_num = MsgBox("是否第一次申请还款", vbYesNo)
    Select Case return_num
    Case 6 'yes
        MsgBox ("不需要选“上一次扣款月份 ！”")
        Month_Choose.Enabled = False
    Case 7   'no
        Month_Choose.Enabled = True
    End Select
Case "合同最后一年度还款"
    '''实际天数到9.20号
    Month_Choose.Enabled = True
    Month_Choose.Clear
    Dim i As Integer
    For i = 1 To 8
    Month_Choose.AddItem (i)
    Next
    Month_Choose.AddItem (12)
Case "既非第一年也非最后一年"
    '''实际天数到9.20号
    Month_Choose.Enabled = True
    For i = 1 To 10
    Month_Choose.AddItem (i)
Next
Month_Choose.AddItem (12)
End Select
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''按照合同规定来还款的计算''''''''''''''''''''''''''''''''''
Private Sub Most_Interest_Click()
'恢复备注信息文本框的显示设置
Interest.ForeColor = vbBlack
Interest.BackColor = vbWhite
Interest.FontSize = 9.5
Interest.FontBold = False
loan_totoal_year_ = Val(Loan_Totoal_Year_Text.Text)
If Most_Interest.Caption = "按规定还款利息计算" Then
    Dim i, j As Integer
    j = 0
    For i = 1 To loan_totoal_year_
        Dim interest_temp As Double
        If i = 1 Then
            Interest.Text = "第" & i & "年利息=" & VBA.Format(Val(Total_Loans_text.Text) * _
                            Val(Interest_Rate_Text.Text) / 100 * 110 / 365, "#0.0") & "元      "
            interest_temp = VBA.Format(Val(Total_Loans_text.Text) * Val(Interest_Rate_Text.Text) _
                            / 100 * 110 / 365, "#0.0")
        ElseIf (i > 1 And i <= Val(No_Principal_Year.Text)) Then
            Interest.Text = Interest.Text & "第" & i & "年利息=" & VBA.Format(Val(Total_Loans_text.Text) _
                            * Val(Interest_Rate_Text.Text) / 100, "#0.0") & "元      "
            interest_temp = VBA.Format(Val(Total_Loans_text.Text) * Val(Interest_Rate_Text.Text) / 100, "#0.0")
        ElseIf (i > Val(No_Principal_Year.Text) And i <= Val(Loan_Totoal_Year_Text.Text) - 1) Then
            Interest.Text = Interest.Text & "第" & i & "年利息=" & VBA.Format((Val(Total_Loans_text.Text) _
                            - j * Val(Common_Year_Principal.Text)) * Val(Interest_Rate_Text.Text) / 100, "#0.0") _
                            & "元      "
            interest_temp = VBA.Format((Val(Total_Loans_text.Text) - j * Val(Common_Year_Principal.Text)) _
                            * Val(Interest_Rate_Text.Text) / 100, "#0.0")
            j = j + 1
        Else
            Interest.Text = Interest.Text & "第" & i & "年利息=" & VBA.Format((Val(Total_Loans_text.Text) _
                            - j * Val(Common_Year_Principal.Text)) * Val(Interest_Rate_Text.Text) / 100, "#0.0") _
                            & "元      "
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
    Most_Interest.Caption = "清除结果重新计算"
ElseIf Most_Interest.Caption = "清除结果重新计算" Then
    Total_Loans_text.Text = 7000
    Interest_Rate_Text.Text = 4.9
    Loan_Totoal_Year_Text.Text = 10
    No_Principal_Year.Text = 3
    Common_Year_Principal.Text = 1000
    Last_Year_Principal.Text = 1000
    Interest.Text = ""
    Most_Interest.Caption = "按规定还款利息计算"
End If


End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''加载窗体初始化''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
''''''''''''''''''''''一些初始化''''''''''''''''''

Dim i As Integer '用于循环的参数

Date.Caption = Format(Now, "yyyy-mm-dd")
Time.Caption = Format(Now, "hh:mm:ss")
'提取系统时间
month_now = Val(Format(Now, "mm"))
day_now = Val(Format(Now, "dd"))
'填充loan_year combobox内容  选择还款年份

Loan_Year.AddItem ("合同第一年度还款")
Loan_Year.AddItem ("合同最后一年度还款")
Loan_Year.AddItem ("既非第一年也非最后一年")
'填充Month_Choose combobox内容    选择上一次还款月份
Month_Choose.AddItem (12)
For i = 1 To 11
    Month_Choose.AddItem (i)
Next

'填充Today_Request combobox内容    选择是否今天申请还款

Today_Request.AddItem ("是")
Today_Request.AddItem ("否")
'填充Predict_Month combobox内容    选择还款月份
Predict_Month.Clear
For i = month_now To 11
    Predict_Month.AddItem (i)
Next
'填充Predict_Day combobox内容    选择还款日期
Predict_Day.Clear
For i = 1 To 31
    Predict_Day.AddItem (i)
Next
 
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''打开Github链接''''''''''''''''''''''''''''''''''''''
Private Sub github_Click()
return_num = ShellExecute(Me.hwnd, "open", "https://github.com/geekerboy/Student_Loans_Calculator.git", _
"", "", 1)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''提前还款的计算''''''''''''''''''''''''''''''''''''''
Private Sub Repay_Advance_Click()
  ''''''''''''''''''''''提前还款相关参数初始化'''''''''''''''
Dim principal_now As Integer
Dim repay_year As Integer
Dim rate_now As Double
Dim reduction_month As Integer
Dim repay_pricipal As Integer
Dim repay_today As Boolean
'Dim month_reall As Integer  '实际参与利息天数计算的月份变量
'Dim day_reall As Integer  '实际参与利息天数计算的日期变量
Dim total_day As Integer  '参与利息计算的天数
Dim repay_interest As Double  '提前还款的利息部分
Dim month_left As Long '申请还款后离今年结束还需要还款的月份天数
Dim year_last_month As String  '本年度结束的最后月份
Dim year_str As String         '显示特殊年度
Dim add_info As String         '显示距离本年度结束还有多少利息
If Repay_Advance.Caption = "计算提前还款利息" Then
    
    '''''利息=本金*利率*天数
    ''''''''''''''''''''''''''''''''''''''参数更新'''''''''''''''''''''''''''''''
    principal_now = Val(Remaining_Principle_Text.Text)    '''1.获取当前剩余贷款本金
    
    rate_now = Val(Rate_Text.Text)           '''''2.获取利率
    reduction_month = Val(Month_Choose.Text)
    repay_pricipal = Val(Reapy_Principal_Text.Text)
    Select Case Today_Request.Text
    Case "是"
        month_reall = month_now
        day_reall = day_now
    Case "否"
        month_reall = Val(Predict_Month.Text)
        day_reall = Val(Predict_Day.Text)
    End Select
    
        If month_reall = 11 Then
            month_reall = 12
        End If
        If month_reall = 10 And day_reall > 15 Then
            month_reall = 12
        End If
        If month_reall = 10 And day_reall < 16 Then
            month_reall = month_reall + 1
            'MsgBox (day_reall)
        End If
    
        If day_reall > 15 And month_reall < 10 Then
            month_reall = month_reall + 1
            'MsgBox (123)
        End If
    
    day_reall = 21
    
    year_last_month = "12.21"
    Select Case Loan_Year.Text
    Case "合同第一年度还款"
        '''实际天数为9.1到12.20号
       year_str = "这是第一年度还款" & vbCrLf
        month_left = 12 - month_reall
    Case "合同最后一年度还款"
        '''实际天数为365
        year_str = "这是最后一年度还款" & vbCrLf
        month_left = 9 - month_reall
        year_last_month = "9.21"
    Case "既非第一年也非最后一年"
        '''实际天数到9.20号
        month_left = 12 - month_reall
    End Select
    If reduction_month = 12 Then
        total_day = month_reall * 30
    Else
        total_day = (month_reall - reduction_month) * 30
    End If
    If month_left = 0 Then
        add_info = "本年度利息结算完毕，本年度不再产生利息"
    Else
        add_info = "如不再申请本年度到" & year_last_month & "号将还利息" & VBA.Format(month_left * _
                                30 * (principal_now - repay_pricipal) * rate_now / 365 / 100, "#0.0") & "元"
    End If
    '''''''''''''''''''''''''''''真正计算''''''''''''''''''''''''''''
    Result_Repay_Advance.FontSize = 9.5
    repay_interest = VBA.Format(rate_now * principal_now * total_day / 365 / 100, "#0.0")
    Result_Repay_Advance.Text = year_str & "将在" & month_reall & "月" & day_reall & "日扣款" & _
                                repay_pricipal + repay_interest & "元," & "其中利息" & repay_interest & _
                                  "元" & vbCrLf & add_info '_
                                '"如不再申请本年度到" & year_last_month & "号将还利息" & VBA.Format(month_left * _
                                30 * (principal_now - repay_pricipal) * rate_now / 365 / 100, "#0.0") & "元"
    Repay_Advance.Caption = "清除结果重新计算"
ElseIf Repay_Advance.Caption = "清除结果重新计算" Then
    Result_Repay_Advance.Text = ""
    Repay_Advance.Caption = "计算提前还款利息"
End If


End Sub
'''''''''''''''''''''''''''''''''''''''''''''''利息计算的一些说明'''''''''''''''''''''''''''''''''
Private Sub Settlement_Cycle_Click(Index As Integer)
    Select Case Index
    Case 1
       msg_return = MsgBox("本计算针对的是生源地贷款的利息计算" & vbCrLf & _
                "1.从毕业年份的9月1号开始计算利息，直到合同最后一年的9月20日结束" & vbCrLf & _
                "2.生源地贷款每月15日（含）前申请提前还款的，当月21号扣款，利息计算天数截止" & vbCrLf & _
                    "   当月20号" & vbCrLf & _
                "3.16日(含）以后申请的，下一个月的21号扣款，利息计算天数截止下一个月的20号" & vbCrLf & _
                "4.10.16-11.30日期间申请的还款，就相当于是申请结算当年剩余需要还的利息（也就是" _
                    & vbCrLf & "   截止12.20号的天数）" & vbCrLf & vbCrLf & _
                    "                   更多消息 ，见Github介绍" _
                    & vbCrLf & "                   点击“确定”，直达Github主页" _
                    & vbCrLf & "                   点击“确定”，直达Github主页" _
                    & vbCrLf & "                   点击“确定”，直达Github主页" _
                    & vbCrLf & "                   不需要查看Github,点击“取消”即可", vbOKCancel)
                    
    If msg_return = 1 Then
        return_num = ShellExecute(Me.hwnd, "open", "https://github.com/geekerboy/Student_Loans_Calculator.git", _
"", "", 1)
    End If
    
    End Select
End Sub



Private Sub Timer1_Timer()
    Date.Caption = Format(Now, "yyyy-mm-dd")
    Time.Caption = Format(Now, "hh:mm:ss")
End Sub

Private Sub Today_Request_Click()
Select Case Today_Request.Text
Case "是"
    Predict_date_.Visible = False
    Predict_Month.Visible = False
    month(0).Visible = False
    Predict_Day.Visible = False
    day.Visible = False
Case "否"
    Predict_date_.Visible = True
    Predict_Month.Visible = True
    month(0).Visible = True
    Predict_Day.Visible = True
    day.Visible = True
End Select
End Sub


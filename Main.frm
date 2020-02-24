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
   StartUpPosition =   3  '����ȱʡ
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
      Caption         =   "������ǰ������Ϣ"
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
      Text            =   "��"
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
      Text            =   "�ȷǵ�һ��Ҳ�����һ��"
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
         Name            =   "����"
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
      ToolTipText     =   "��������ܶ�"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Most_Interest 
      Caption         =   "���涨������Ϣ����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "��"
      Height          =   180
      Index           =   1
      Left            =   7440
      TabIndex        =   42
      Top             =   1680
      Width           =   180
   End
   Begin VB.Label yuan 
      AutoSize        =   -1  'True
      Caption         =   "Ԫ"
      Height          =   180
      Index           =   4
      Left            =   7440
      TabIndex        =   41
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label year 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Index           =   1
      Left            =   3120
      TabIndex        =   40
      Top             =   1680
      Width           =   180
   End
   Begin VB.Label year 
      AutoSize        =   -1  'True
      Caption         =   "��"
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
      Caption         =   "Ԫ"
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
      Caption         =   "Ԫ"
      Height          =   180
      Index           =   2
      Left            =   3120
      TabIndex        =   34
      Top             =   2640
      Width           =   180
   End
   Begin VB.Label yuan 
      AutoSize        =   -1  'True
      Caption         =   "Ԫ"
      Height          =   180
      Index           =   1
      Left            =   3120
      TabIndex        =   33
      Top             =   2160
      Width           =   180
   End
   Begin VB.Label yuan 
      AutoSize        =   -1  'True
      Caption         =   "Ԫ"
      Height          =   180
      Index           =   0
      Left            =   3120
      TabIndex        =   32
      Top             =   240
      Width           =   180
   End
   Begin VB.Label day 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   7920
      TabIndex        =   28
      Top             =   3000
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label month 
      AutoSize        =   -1  'True
      Caption         =   "��"
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
      Caption         =   "Ԥ������黹����"
      Height          =   180
      Left            =   4200
      TabIndex        =   25
      Top             =   3000
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Today_Request_ 
      AutoSize        =   -1  'True
      Caption         =   "�Ƿ�������뻹��"
      Height          =   180
      Left            =   4200
      TabIndex        =   24
      Top             =   2520
      Width           =   1440
   End
   Begin VB.Label Reapy_Principal_ 
      AutoSize        =   -1  'True
      Caption         =   "����黹����"
      Height          =   180
      Left            =   4320
      TabIndex        =   23
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Label Rate_ 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   4800
      TabIndex        =   22
      Top             =   1200
      Width           =   360
   End
   Begin VB.Label Last_Deduction 
      AutoSize        =   -1  'True
      Caption         =   "��һ�οۿ��·�"
      Height          =   180
      Left            =   4320
      TabIndex        =   17
      Top             =   1560
      Width           =   1260
   End
   Begin VB.Label Loan_Yeear_ 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      Height          =   180
      Left            =   4680
      TabIndex        =   16
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Remaining_Principle_ 
      AutoSize        =   -1  'True
      Caption         =   "��ǰʣ�౾��"
      Height          =   180
      Left            =   4560
      TabIndex        =   15
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label Loan_Totoal_Year 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Left            =   600
      TabIndex        =   13
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Interest_Rate 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   840
      TabIndex        =   12
      Top             =   720
      Width           =   360
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "���һ�껹����"
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "�м�ÿ�껹������"
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "���û���������"
      Height          =   180
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label Total_Loans 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "�����ܶ�"
      Height          =   180
      Left            =   720
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.Menu Deep_Calculate 
      Caption         =   "��ȼ���"
      Index           =   0
      Begin VB.Menu Fixed_Principal 
         Caption         =   "ÿ�껹�����������"
         Index           =   1
      End
   End
   Begin VB.Menu Calculate_Info 
      Caption         =   "����˵��"
      Index           =   0
      Begin VB.Menu Settlement_Cycle 
         Caption         =   "��������˵��"
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
'�ò˵������Ϣ��������˵��
'��ӵ�λ������һ���Ű�ϸ��
'���ʱ����ʾ

Option Explicit '�������붨�����ʹ��

'ShellExecute������Ҫ�����������
''''��һ����ShellExecute������Ҫ�����һЩ����
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'''''ShellExecute�����������ļ����������ӵ�ɧ����

Dim return_num As Integer  '����ShellExecute����������ֵ��û�з���ֵ���ջᱨ��
Dim loan_totoal_year_ As Integer
'���Դ�http://note.youdao.com/noteshare?id=3bfdcedbb35f23c29d5fbc1a4f0c8a58
'return_num = ShellExecute(Me.hwnd, "open", "http://note.youdao.com/noteshare?id=3bfdcedbb35f23c29d5fbc1a4f0c8a58 ", "", "", 1)
Dim total_interest As Double  '��Ϣ��ͳ��
Dim month_now As Integer '��ȡϵͳ��ǰ�·�
Dim day_now As Integer '��ȡϵͳ��ǰ���������
Dim month_reall As Integer  'ʵ�ʲ�����Ϣ����������·ݱ���
Dim day_reall As Integer  'ʵ�ʲ�����Ϣ������������ڱ���
Dim msg_return As Integer 'msg��Ϣ����ֵ


Private Sub Loan_Year_Click()
Select Case Loan_Year.Text
Case "��ͬ��һ��Ȼ���"
    '''ʵ������Ϊ9.1��12.20��
    Month_Choose.Clear
    Month_Choose.AddItem (9)
    Month_Choose.AddItem (10)
    return_num = MsgBox("�Ƿ��һ�����뻹��", vbYesNo)
    Select Case return_num
    Case 6 'yes
        MsgBox ("����Ҫѡ����һ�οۿ��·� ����")
        Month_Choose.Enabled = False
    Case 7   'no
        Month_Choose.Enabled = True
    End Select
Case "��ͬ���һ��Ȼ���"
    '''ʵ��������9.20��
    Month_Choose.Enabled = True
    Month_Choose.Clear
    Dim i As Integer
    For i = 1 To 8
    Month_Choose.AddItem (i)
    Next
    Month_Choose.AddItem (12)
Case "�ȷǵ�һ��Ҳ�����һ��"
    '''ʵ��������9.20��
    Month_Choose.Enabled = True
    For i = 1 To 10
    Month_Choose.AddItem (i)
Next
Month_Choose.AddItem (12)
End Select
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''���պ�ͬ�涨������ļ���''''''''''''''''''''''''''''''''''
Private Sub Most_Interest_Click()
'�ָ���ע��Ϣ�ı������ʾ����
Interest.ForeColor = vbBlack
Interest.BackColor = vbWhite
Interest.FontSize = 9.5
Interest.FontBold = False
loan_totoal_year_ = Val(Loan_Totoal_Year_Text.Text)
If Most_Interest.Caption = "���涨������Ϣ����" Then
    Dim i, j As Integer
    j = 0
    For i = 1 To loan_totoal_year_
        Dim interest_temp As Double
        If i = 1 Then
            Interest.Text = "��" & i & "����Ϣ=" & VBA.Format(Val(Total_Loans_text.Text) * _
                            Val(Interest_Rate_Text.Text) / 100 * 110 / 365, "#0.0") & "Ԫ      "
            interest_temp = VBA.Format(Val(Total_Loans_text.Text) * Val(Interest_Rate_Text.Text) _
                            / 100 * 110 / 365, "#0.0")
        ElseIf (i > 1 And i <= Val(No_Principal_Year.Text)) Then
            Interest.Text = Interest.Text & "��" & i & "����Ϣ=" & VBA.Format(Val(Total_Loans_text.Text) _
                            * Val(Interest_Rate_Text.Text) / 100, "#0.0") & "Ԫ      "
            interest_temp = VBA.Format(Val(Total_Loans_text.Text) * Val(Interest_Rate_Text.Text) / 100, "#0.0")
        ElseIf (i > Val(No_Principal_Year.Text) And i <= Val(Loan_Totoal_Year_Text.Text) - 1) Then
            Interest.Text = Interest.Text & "��" & i & "����Ϣ=" & VBA.Format((Val(Total_Loans_text.Text) _
                            - j * Val(Common_Year_Principal.Text)) * Val(Interest_Rate_Text.Text) / 100, "#0.0") _
                            & "Ԫ      "
            interest_temp = VBA.Format((Val(Total_Loans_text.Text) - j * Val(Common_Year_Principal.Text)) _
                            * Val(Interest_Rate_Text.Text) / 100, "#0.0")
            j = j + 1
        Else
            Interest.Text = Interest.Text & "��" & i & "����Ϣ=" & VBA.Format((Val(Total_Loans_text.Text) _
                            - j * Val(Common_Year_Principal.Text)) * Val(Interest_Rate_Text.Text) / 100, "#0.0") _
                            & "Ԫ      "
            interest_temp = VBA.Format((Val(Total_Loans_text.Text) - j * Val(Common_Year_Principal.Text)) _
                            * Val(Interest_Rate_Text.Text) / 100, "#0.0")
        End If
        
        total_interest = total_interest + interest_temp
        
        If (i Mod 2 = 0) And i <> loan_totoal_year_ Then
                Interest.Text = Interest.Text & vbCrLf
        End If
    Next i
    Interest.Text = Interest.Text & vbCrLf & "��������Ϣ=" & total_interest & "Ԫ"
    total_interest = 0
    Most_Interest.Caption = "���������¼���"
ElseIf Most_Interest.Caption = "���������¼���" Then
    Total_Loans_text.Text = 7000
    Interest_Rate_Text.Text = 4.9
    Loan_Totoal_Year_Text.Text = 10
    No_Principal_Year.Text = 3
    Common_Year_Principal.Text = 1000
    Last_Year_Principal.Text = 1000
    Interest.Text = ""
    Most_Interest.Caption = "���涨������Ϣ����"
End If


End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���ش����ʼ��''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
''''''''''''''''''''''һЩ��ʼ��''''''''''''''''''

Dim i As Integer '����ѭ���Ĳ���

Date.Caption = Format(Now, "yyyy-mm-dd")
Time.Caption = Format(Now, "hh:mm:ss")
'��ȡϵͳʱ��
month_now = Val(Format(Now, "mm"))
day_now = Val(Format(Now, "dd"))
'���loan_year combobox����  ѡ�񻹿����

Loan_Year.AddItem ("��ͬ��һ��Ȼ���")
Loan_Year.AddItem ("��ͬ���һ��Ȼ���")
Loan_Year.AddItem ("�ȷǵ�һ��Ҳ�����һ��")
'���Month_Choose combobox����    ѡ����һ�λ����·�
Month_Choose.AddItem (12)
For i = 1 To 11
    Month_Choose.AddItem (i)
Next

'���Today_Request combobox����    ѡ���Ƿ�������뻹��

Today_Request.AddItem ("��")
Today_Request.AddItem ("��")
'���Predict_Month combobox����    ѡ�񻹿��·�
Predict_Month.Clear
For i = month_now To 11
    Predict_Month.AddItem (i)
Next
'���Predict_Day combobox����    ѡ�񻹿�����
Predict_Day.Clear
For i = 1 To 31
    Predict_Day.AddItem (i)
Next
 
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''��Github����''''''''''''''''''''''''''''''''''''''
Private Sub github_Click()
return_num = ShellExecute(Me.hwnd, "open", "https://github.com/geekerboy/Student_Loans_Calculator.git", _
"", "", 1)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''��ǰ����ļ���''''''''''''''''''''''''''''''''''''''
Private Sub Repay_Advance_Click()
  ''''''''''''''''''''''��ǰ������ز�����ʼ��'''''''''''''''
Dim principal_now As Integer
Dim repay_year As Integer
Dim rate_now As Double
Dim reduction_month As Integer
Dim repay_pricipal As Integer
Dim repay_today As Boolean
'Dim month_reall As Integer  'ʵ�ʲ�����Ϣ����������·ݱ���
'Dim day_reall As Integer  'ʵ�ʲ�����Ϣ������������ڱ���
Dim total_day As Integer  '������Ϣ���������
Dim repay_interest As Double  '��ǰ�������Ϣ����
Dim month_left As Long '���뻹���������������Ҫ������·�����
Dim year_last_month As String  '����Ƚ���������·�
Dim year_str As String         '��ʾ�������
Dim add_info As String         '��ʾ���뱾��Ƚ������ж�����Ϣ
If Repay_Advance.Caption = "������ǰ������Ϣ" Then
    
    '''''��Ϣ=����*����*����
    ''''''''''''''''''''''''''''''''''''''��������'''''''''''''''''''''''''''''''
    principal_now = Val(Remaining_Principle_Text.Text)    '''1.��ȡ��ǰʣ������
    
    rate_now = Val(Rate_Text.Text)           '''''2.��ȡ����
    reduction_month = Val(Month_Choose.Text)
    repay_pricipal = Val(Reapy_Principal_Text.Text)
    Select Case Today_Request.Text
    Case "��"
        month_reall = month_now
        day_reall = day_now
    Case "��"
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
    Case "��ͬ��һ��Ȼ���"
        '''ʵ������Ϊ9.1��12.20��
       year_str = "���ǵ�һ��Ȼ���" & vbCrLf
        month_left = 12 - month_reall
    Case "��ͬ���һ��Ȼ���"
        '''ʵ������Ϊ365
        year_str = "�������һ��Ȼ���" & vbCrLf
        month_left = 9 - month_reall
        year_last_month = "9.21"
    Case "�ȷǵ�һ��Ҳ�����һ��"
        '''ʵ��������9.20��
        month_left = 12 - month_reall
    End Select
    If reduction_month = 12 Then
        total_day = month_reall * 30
    Else
        total_day = (month_reall - reduction_month) * 30
    End If
    If month_left = 0 Then
        add_info = "�������Ϣ������ϣ�����Ȳ��ٲ�����Ϣ"
    Else
        add_info = "�粻�����뱾��ȵ�" & year_last_month & "�Ž�����Ϣ" & VBA.Format(month_left * _
                                30 * (principal_now - repay_pricipal) * rate_now / 365 / 100, "#0.0") & "Ԫ"
    End If
    '''''''''''''''''''''''''''''��������''''''''''''''''''''''''''''
    Result_Repay_Advance.FontSize = 9.5
    repay_interest = VBA.Format(rate_now * principal_now * total_day / 365 / 100, "#0.0")
    Result_Repay_Advance.Text = year_str & "����" & month_reall & "��" & day_reall & "�տۿ�" & _
                                repay_pricipal + repay_interest & "Ԫ," & "������Ϣ" & repay_interest & _
                                  "Ԫ" & vbCrLf & add_info '_
                                '"�粻�����뱾��ȵ�" & year_last_month & "�Ž�����Ϣ" & VBA.Format(month_left * _
                                30 * (principal_now - repay_pricipal) * rate_now / 365 / 100, "#0.0") & "Ԫ"
    Repay_Advance.Caption = "���������¼���"
ElseIf Repay_Advance.Caption = "���������¼���" Then
    Result_Repay_Advance.Text = ""
    Repay_Advance.Caption = "������ǰ������Ϣ"
End If


End Sub
'''''''''''''''''''''''''''''''''''''''''''''''��Ϣ�����һЩ˵��'''''''''''''''''''''''''''''''''
Private Sub Settlement_Cycle_Click(Index As Integer)
    Select Case Index
    Case 1
       msg_return = MsgBox("��������Ե�����Դ�ش������Ϣ����" & vbCrLf & _
                "1.�ӱ�ҵ��ݵ�9��1�ſ�ʼ������Ϣ��ֱ����ͬ���һ���9��20�ս���" & vbCrLf & _
                "2.��Դ�ش���ÿ��15�գ�����ǰ������ǰ����ģ�����21�ſۿ��Ϣ����������ֹ" & vbCrLf & _
                    "   ����20��" & vbCrLf & _
                "3.16��(�����Ժ�����ģ���һ���µ�21�ſۿ��Ϣ����������ֹ��һ���µ�20��" & vbCrLf & _
                "4.10.16-11.30���ڼ�����Ļ�����൱����������㵱��ʣ����Ҫ������Ϣ��Ҳ����" _
                    & vbCrLf & "   ��ֹ12.20�ŵ�������" & vbCrLf & vbCrLf & _
                    "                   ������Ϣ ����Github����" _
                    & vbCrLf & "                   �����ȷ������ֱ��Github��ҳ" _
                    & vbCrLf & "                   �����ȷ������ֱ��Github��ҳ" _
                    & vbCrLf & "                   �����ȷ������ֱ��Github��ҳ" _
                    & vbCrLf & "                   ����Ҫ�鿴Github,�����ȡ��������", vbOKCancel)
                    
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
Case "��"
    Predict_date_.Visible = False
    Predict_Month.Visible = False
    month(0).Visible = False
    Predict_Day.Visible = False
    day.Visible = False
Case "��"
    Predict_date_.Visible = True
    Predict_Month.Visible = True
    month(0).Visible = True
    Predict_Day.Visible = True
    day.Visible = True
End Select
End Sub


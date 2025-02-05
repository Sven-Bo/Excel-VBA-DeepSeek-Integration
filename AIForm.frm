VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AIForm 
   Caption         =   "AI助手"
   ClientHeight    =   11160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8055
   OleObjectBlob   =   "AIForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "AIForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    #If VBA7 And Win64 Then
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
        Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    #Else
        Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
        Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
        Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    #End If
    
    Dim model As String
    
Private Sub QuestionLabel_Click()
    QuestionLabel.Visible = False
    QuestionTextbox.SetFocus
End Sub

Private Sub QuestionTextbox_Change()
    QuestionLabel.Visible = False
End Sub

Private Sub QuestionTextbox_Enter()
    QuestionLabel.Visible = False
End Sub

Private Sub SendButton_Click()
    Call SetApiKey '设置api key

    Dim question As String, allMessage As String

    '1 将问题发送出去 等待ai回答
    question = QuestionTextbox.Text
    QuestionTextbox.Text = ""
    If question = "" Then
        Exit Sub
    End If
    
    
    '2 显示问题
    allMessage = ChatTextBox.Text
    allMessage = allMessage & "我：" & question & vbCrLf & vbCrLf
    
    ChatTextBox.Text = allMessage
    
    WaitingLabel.Visible = True
    DoEvents
    
    '3 显示ai回答
    Dim answer As String
    answer = mDeepSeek.DS_Chat(question)
    
    allMessage = ChatTextBox.Text
    allMessage = allMessage & "DeepSeek：" & answer & vbCrLf & vbCrLf
    
    ChatTextBox.Text = allMessage
    
    WaitingLabel.Visible = False
End Sub

Sub SetApiKey()
    Dim apiKey As String
    
    If R1ModelCheckBox.Enabled Then
        '深度思考
        model = "deepseek-reasoner"
    Else
        model = "deepseek-chat"
    End If
    
    '需填写自己的api Key
    apiKey = ""
    
    '设置上key和使用的模型
    Initial.InitialDeepSeekKey apiKey, model
End Sub
    
    '窗体初始化时设置
Private Sub UserForm_Initialize()
    WaitingLabel.Visible = False
    
    ChatTextBox.MultiLine = True
    ChatTextBox.WordWrap = True
    ChatTextBox.EnterKeyBehavior = True
    
    QuestionTextbox.MultiLine = True
    QuestionTextbox.WordWrap = True
    QuestionTextbox.EnterKeyBehavior = True
    
    '添加最大最小化按钮
    SetWindowLong FindWindow("ThunderDFrame", Me.Caption), -16, GetWindowLong(FindWindow("ThunderDFrame", Me.Caption), -16) Or &H40000 Or &H20000 Or &H10000
    
    SendButton.SetFocus
End Sub

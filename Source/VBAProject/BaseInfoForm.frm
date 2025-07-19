VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BaseInfoForm 
   Caption         =   "论文基础信息"
   ClientHeight    =   6045
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8010
   OleObjectBlob   =   "BaseInfoForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "BaseInfoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub OkBtn_Click() ' 确定按钮
    Dim ur  As UndoRecord

    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "更新基础信息"

    titleCN = tbTitleCN.Value
    titleEN = tbTitleEN.Value
    studentName = tbName.Value
    studentNo = tbStudentNo.Value
    teacherName = tbTeacherName.Value
    teacherTitle = tbTeacherTitle.Value
    major = tbMajor.Value

    UpdateContentControl "封面题目", Trim(titleCN)
    UpdateContentControl "论文题目", Trim(titleCN)
    UpdateContentControl "英文题目", Trim(titleEN)
    UpdateContentControl "作者", Trim(studentName)
    UpdateContentControl "编号", Trim(studentNo)
    UpdateContentControl "导师", Trim(teacherName)
    UpdateContentControl "职称", Trim(teacherTitle)
    UpdateContentControl "专业", Trim(major)

    Unload Me ' 关闭窗体

    ur.EndCustomRecord
    Exit Sub

ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "更新基础信息时发生错误: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Private Sub CancelBtn_Click() ' 取消按钮
    Unload Me
End Sub

Private Function GetContentControl(title As String) As String
    Dim cc As ContentControl
    
    ' 通过标题(Title)查找并更新内容控件
    On Error Resume Next
    Set cc = ActiveDocument.SelectContentControlsByTitle(title).item(1)
    On Error GoTo 0
    GetContentControl = cc.Range.text
End Function

Private Sub UpdateContentControl(title As String, val As String)
    Dim cc As ContentControl
    
    ' 通过标题(Title)查找并更新内容控件
    On Error Resume Next
    Set cc = ActiveDocument.SelectContentControlsByTitle(title).item(1)
    On Error GoTo 0
    
    If Not cc Is Nothing Then
        ' 或者使用以下方式设置纯文本内容控件的值
        cc.LockContents = False ' 先解锁(如果需要)
        cc.Range.text = val
        cc.LockContents = True ' 重新锁定(如果需要)
        
        'MsgBox "内容控件已更新!", vbInformation, C_TITLE
    Else
        MsgBox "未找到指定标题的内容控件!", vbExclamation, C_TITLE
    End If
End Sub


Private Sub UserForm_Initialize()
    titleCN = GetContentControl("论文题目")
    titleEN = GetContentControl("英文题目")
    studentName = GetContentControl("作者")
    studentNo = GetContentControl("编号")
    teacherName = GetContentControl("导师")
    teacherTitle = GetContentControl("职称")
    major = GetContentControl("专业")
    
    If titleCN <> "论文中文题目" Then
        tbTitleCN.Value = titleCN
    End If
    
    If titleEN <> "Thesis Title" Then
        tbTitleEN.Value = titleEN
    End If
    
    If studentName <> "作者姓名" Then
        tbName.Value = studentName
    End If
    
    If studentNo <> "学号" Then
        tbStudentNo.Value = studentNo
    End If
    
    If teacherName <> "导师姓名" Then
        tbTeacherName.Value = teacherName
    End If
    
    If teacherTitle <> "职称" Then
        tbTeacherTitle.Value = teacherTitle
    End If
    
    If major <> "专业" Then
        tbMajor.Value = major
    End If
    
End Sub


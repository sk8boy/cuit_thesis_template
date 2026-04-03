VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BaseInfoForm 
   Caption         =   "论文基础信息"
   ClientHeight    =   9015.001
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
    college = tbCollege.Value
    major = tbMajor.Value
    studyField = tbStudyField.Value
    firstTeacherName = tbFirstTeacherName.Value
    firstTeacherTitle = tbFirstTeacherTitle.Value
    otherTeacherName = tbOtherTeacherName.Value
    otherTeacherTitle = tbOtherTeacherTitle.Value

    UpdateContentControl "封面中文题目", Trim(titleCN)
    UpdateContentControl "论文题目", Trim(titleCN)

    UpdateContentControl "封面英文题目", Trim(titleEN)

    UpdateContentControl "作者", Trim(studentName)
    
    UpdateContentControl "学号", Trim(studentNo)
    
    UpdateContentControl "学院", Trim(college)
    
    UpdateContentControl "专业", Trim(major)
    
    UpdateContentControl "研究方向", Trim(studyField)
    
    UpdateContentControl "导师", Trim(firstTeacherName)
    UpdateContentControl "指导老师", Trim(firstTeacherName)

    UpdateContentControl "职称", Trim(firstTeacherTitle)

    UpdateContentControl "辅助导师", Trim(otherTeacherName)

    UpdateContentControl "辅助导师职称", Trim(otherTeacherTitle)

    Unload Me ' 关闭窗体

    ur.EndCustomRecord
    Exit Sub

ERROR_HANDLER:
    If err.Number = ERR_USRMSG Then
        MsgBox err.Description, vbExclamation, C_TITLE
    ElseIf err.Number <> ERR_CANCEL Then
        MsgBox "更新基础信息时发生错误: " & err.Description, vbCritical, C_TITLE
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
    titleCN = GetContentControl("封面中文题目")
    titleEN = GetContentControl("封面英文题目")
    studentName = GetContentControl("作者")
    studentNo = GetContentControl("学号")
    college = GetContentControl("学院")
    major = GetContentControl("专业")
    studyField = GetContentControl("研究方向")
    firstTeacherName = GetContentControl("导师")
    firstTeacherTitle = GetContentControl("职称")
    otherTeacherName = GetContentControl("辅助导师")
    otherTeacherTitle = GetContentControl("辅助导师职称")
    
    If titleCN <> "论文中文题目" Then
        tbTitleCN.Value = titleCN
    End If
    
    If titleEN <> "Thesis Title" Then
        tbTitleEN.Value = titleEN
    End If
    
    If studentName <> "姓名" Then
        tbName.Value = studentName
    End If
    
    If studentNo <> "学号" Then
        tbStudentNo.Value = studentNo
    End If
    
    If college <> "学院" Then
        tbCollege.Value = college
    End If
    
    If major <> "专业" Then
        tbMajor.Value = major
    End If
    
    If studyField <> "研究方向" Then
        tbStudyField.Value = studyField
    End If
    
    If firstTeacherName <> "导师姓名" Then
        tbFirstTeacherName.Value = firstTeacherName
    End If
    
    If firstTeacherTitle <> "职称" Then
        tbFirstTeacherTitle.Value = firstTeacherTitle
    End If
    
    If otherTeacherName <> "导师姓名" Then
        tbOtherTeacherName.Value = otherTeacherName
    End If
    
    If otherTeacherTitle <> "职称" Then
        tbOtherTeacherTitle.Value = otherTeacherTitle
    End If
End Sub


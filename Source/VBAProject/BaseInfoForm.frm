VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BaseInfoForm 
   Caption         =   "���Ļ�����Ϣ"
   ClientHeight    =   6045
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8010
   OleObjectBlob   =   "BaseInfoForm.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "BaseInfoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub OkBtn_Click() ' ȷ����ť
    Dim ur  As UndoRecord

    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "���»�����Ϣ"

    titleCN = tbTitleCN.Value
    titleEN = tbTitleEN.Value
    studentName = tbName.Value
    studentNo = tbStudentNo.Value
    teacherName = tbTeacherName.Value
    teacherTitle = tbTeacherTitle.Value
    major = tbMajor.Value

    UpdateContentControl "������Ŀ", Trim(titleCN)
    UpdateContentControl "������Ŀ", Trim(titleCN)
    UpdateContentControl "Ӣ����Ŀ", Trim(titleEN)
    UpdateContentControl "����", Trim(studentName)
    UpdateContentControl "���", Trim(studentNo)
    UpdateContentControl "��ʦ", Trim(teacherName)
    UpdateContentControl "ְ��", Trim(teacherTitle)
    UpdateContentControl "רҵ", Trim(major)

    Unload Me ' �رմ���

    ur.EndCustomRecord
    Exit Sub

ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "���»�����Ϣʱ��������: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Private Sub CancelBtn_Click() ' ȡ����ť
    Unload Me
End Sub

Private Function GetContentControl(title As String) As String
    Dim cc As ContentControl
    
    ' ͨ������(Title)���Ҳ��������ݿؼ�
    On Error Resume Next
    Set cc = ActiveDocument.SelectContentControlsByTitle(title).item(1)
    On Error GoTo 0
    GetContentControl = cc.Range.text
End Function

Private Sub UpdateContentControl(title As String, val As String)
    Dim cc As ContentControl
    
    ' ͨ������(Title)���Ҳ��������ݿؼ�
    On Error Resume Next
    Set cc = ActiveDocument.SelectContentControlsByTitle(title).item(1)
    On Error GoTo 0
    
    If Not cc Is Nothing Then
        ' ����ʹ�����·�ʽ���ô��ı����ݿؼ���ֵ
        cc.LockContents = False ' �Ƚ���(�����Ҫ)
        cc.Range.text = val
        cc.LockContents = True ' ��������(�����Ҫ)
        
        'MsgBox "���ݿؼ��Ѹ���!", vbInformation, C_TITLE
    Else
        MsgBox "δ�ҵ�ָ����������ݿؼ�!", vbExclamation, C_TITLE
    End If
End Sub


Private Sub UserForm_Initialize()
    titleCN = GetContentControl("������Ŀ")
    titleEN = GetContentControl("Ӣ����Ŀ")
    studentName = GetContentControl("����")
    studentNo = GetContentControl("���")
    teacherName = GetContentControl("��ʦ")
    teacherTitle = GetContentControl("ְ��")
    major = GetContentControl("רҵ")
    
    If titleCN <> "����������Ŀ" Then
        tbTitleCN.Value = titleCN
    End If
    
    If titleEN <> "Thesis Title" Then
        tbTitleEN.Value = titleEN
    End If
    
    If studentName <> "��������" Then
        tbName.Value = studentName
    End If
    
    If studentNo <> "ѧ��" Then
        tbStudentNo.Value = studentNo
    End If
    
    If teacherName <> "��ʦ����" Then
        tbTeacherName.Value = teacherName
    End If
    
    If teacherTitle <> "ְ��" Then
        tbTeacherTitle.Value = teacherTitle
    End If
    
    If major <> "רҵ" Then
        tbMajor.Value = major
    End If
    
End Sub


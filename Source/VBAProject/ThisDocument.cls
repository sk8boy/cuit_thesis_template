VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Document_Open()
    Dim ur  As UndoRecord
    Dim objField  As Field

    On Error Resume Next
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "更新所有序号"
    'Update fields
    For Each objField In ActiveDocument.Fields
       If objField.Type = wdFieldSequence Then
          objField.Update
       End If
    Next
    ur.EndCustomRecord
    'If no other change is made (except for updating the fields), allow closing without saving
    ActiveDocument.Saved = True
    'Check for standard page setup
    RestorePageSetup
End Sub

' 如果是文档中的内容控件下拉列表
Private Sub Document_ContentControlOnExit(ByVal ContentControl As ContentControl, Cancel As Boolean)
    ' 检查是否是触发下拉列表
    If ContentControl.title = "论文密级" Then
        Dim selectedValue As String
        selectedValue = ContentControl.Range.text
        Call UpdateControlsBySelection(selectedValue)
    End If
End Sub

' 核心控制函数
Private Sub UpdateControlsBySelection(ByVal selection As String)
    Dim cc As ContentControl
    Dim controlGroup As String
    
    ' 定义控件分组
    Dim editableGroup As Collection
    Dim readonlyGroup As Collection
    
    Set editableGroup = New Collection
    Set readonlyGroup = New Collection

    ' 根据不同选项配置控件权限
    Select Case selection
        Case "公开"
            UpdateContentControl "保密期限", "    "
            UpdateContentControl "解密时间", "年 月 日"
            UpdateContentControl "声明表格密级", "    "
            UpdateContentControl "声明文字密级", "     "
            UpdateContentControl "密级作者签名时间", "年 月 日"
            UpdateContentControl "密级导师签名时间", "年 月 日"
            UpdateContentControl "密级保密办审核时间", "年 月 日"
            Call SetGroupEditable(Array("保密期限", "解密时间", "密级作者签名时间", "密级导师签名时间", "密级保密办审核时间"), False)
            
        Case Else
            UpdateContentControl "声明表格密级", selection
            UpdateContentControl "声明文字密级", selection
            Call SetGroupEditable(Array("保密期限", "解密时间", "密级作者签名时间", "密级导师签名时间", "密级保密办审核时间"), True)
            
    End Select
End Sub

' 批量设置控件组
Private Sub SetGroupEditable(ByVal titles As Variant, ByVal editable As Boolean)
    Dim i As Integer
    Dim cc As ContentControl
    
    For i = LBound(titles) To UBound(titles)
        For Each cc In ActiveDocument.ContentControls
            If cc.title = titles(i) Then
                cc.LockContents = Not editable ' 控件内容是否可编辑
                cc.LockContentControl = Not editable ' 控件本身是否可被删除
                Exit For
            End If
        Next cc
    Next i
End Sub

' 获取所有控件标题
Private Function GetAllControlTitles() As Variant
    Dim titles() As String
    Dim cc As ContentControl
    Dim i As Integer
    
    ReDim titles(0 To ActiveDocument.ContentControls.Count - 1)
    i = 0
    
    For Each cc In ActiveDocument.ContentControls
        titles(i) = cc.title
        i = i + 1
    Next cc
    
    GetAllControlTitles = titles
End Function



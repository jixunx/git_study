' ResizePagesToContent.vba
' 对活动 Visio 文档的每一页，执行"适应绘图"操作，
' 使画布（图纸）大小自动收缩/扩展到绘图内容的边界。
'
' 等同于手动操作：全选本页 → 设计 → 适应绘图
'
' 使用方法：
'   1. 打开 Visio，按 Alt+F11 打开 VBA 编辑器
'   2. 插入 > 模块，将此代码粘贴进去
'   3. 运行 ResizeAllPagesToContent 宏

Option Explicit

Public Sub ResizeAllPagesToContent()
    Dim doc As Visio.Document

    If Visio.Documents.Count = 0 Then
        MsgBox "没有打开的 Visio 文档，请先打开一个 .vsdx 文件。", vbExclamation
        Exit Sub
    End If

    Set doc = Visio.ActiveDocument

    Dim page As Visio.Page
    Dim resized As Long
    Dim skipped As Long
    resized = 0
    skipped = 0

    For Each page In doc.Pages
        If page.Shapes.Count > 0 Then
            page.ResizeToFitContents
            resized = resized + 1
        Else
            skipped = skipped + 1
        End If
    Next page

    Dim msg As String
    msg = "完成！共处理 " & resized & " 页。"
    If skipped > 0 Then
        msg = msg & vbCrLf & "（跳过 " & skipped & " 个空页）"
    End If
    MsgBox msg, vbInformation
End Sub

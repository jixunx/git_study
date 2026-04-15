' ExportVisioToSVG.vba
' 将 Visio 文件中每个 Page 导出为 SVG 文件，文件名以 Sheet（Page）名称命名。
' 导出时临时将图纸尺寸收缩到内容边界，使内容充满整个 SVG 图纸；
' 导出完成后通过 Undo Scope 自动回滚，不修改原文件。
'
' 使用方法：
'   1. 打开 Visio，按 Alt+F11 打开 VBA 编辑器
'   2. 插入 > 模块，将此代码粘贴进去
'   3. 修改下方常量（或运行时弹窗选择目录）
'   4. 运行 ExportAllPagesToSVG 宏

Option Explicit

' ---------- 可修改的参数 ----------
' 默认输出目录，留空则运行时弹窗选择
Private Const DEFAULT_OUTPUT_DIR As String = ""

' 内容周围保留的边距（单位：英寸）。0 = 紧贴内容边界；建议值 0.1
Private Const CONTENT_MARGIN_INCH As Double = 0
' ----------------------------------

' 入口：导出当前活动文档的所有 Page 为 SVG
Public Sub ExportAllPagesToSVG()
    Dim doc As Visio.Document
    Dim outputDir As String

    ' 确保有打开的文档
    If Visio.Documents.Count = 0 Then
        MsgBox "没有打开的 Visio 文档，请先打开一个 .vsdx 文件。", vbExclamation
        Exit Sub
    End If

    Set doc = Visio.ActiveDocument

    ' 确定输出目录
    outputDir = Trim(DEFAULT_OUTPUT_DIR)
    If outputDir = "" Then
        outputDir = BrowseForFolder("请选择 SVG 文件的输出目录")
        If outputDir = "" Then
            MsgBox "未选择输出目录，操作已取消。", vbInformation
            Exit Sub
        End If
    End If

    ' 确保路径以路径分隔符结尾
    If Right(outputDir, 1) <> "\" Then outputDir = outputDir & "\"

    ' 导出每个 Page
    Dim exported As Long
    exported = ExportPages(doc, outputDir)

    MsgBox "导出完成！共导出 " & exported & " 个 SVG 文件。" & vbCrLf & _
           "输出目录：" & outputDir, vbInformation
End Sub

' 遍历文档中所有 Page 并逐一导出为 SVG
' 返回成功导出的页面数量
Private Function ExportPages(doc As Visio.Document, outputDir As String) As Long
    Dim page As Visio.Page
    Dim count As Long

    count = 0

    For Each page In doc.Pages
        ' 使用 Page 名称作为文件名，清理非法字符
        Dim safeName As String
        safeName = SanitizeFileName(page.Name)

        ' 若同名文件已存在则加序号避免覆盖
        Dim finalPath As String
        finalPath = UniqueFilePath(outputDir & safeName & ".svg")

        ' 临时收缩图纸到内容边界后导出，导出后自动还原
        ExportPageFitContent page, finalPath

        count = count + 1
    Next page

    ExportPages = count
End Function

' 将图纸临时缩到内容边界后导出为 SVG，完成后通过 Undo Scope 回滚图纸尺寸。
' 原理：Visio SVG 导出以 PageWidth/PageHeight 为 viewBox；
'       ResizeToFitContents 将图纸收缩至所有形状的外接矩形，
'       使导出的 SVG 中内容充满整张图纸。
Private Sub ExportPageFitContent(page As Visio.Page, svgPath As String)
    ' 若页面没有任何形状，直接导出即可
    If page.Shapes.Count = 0 Then
        page.Export svgPath
        Exit Sub
    End If

    ' 开启一个 Undo Scope，所有在此范围内对文档的修改都可一键回滚
    Dim scopeID As Long
    scopeID = Application.BeginUndoScope("TempResizeForSVGExport")

    On Error GoTo ErrHandler

    ' 将图纸尺寸收缩到内容边界（加上可选边距）
    page.ResizeToFitContents CONTENT_MARGIN_INCH

    ' 导出 SVG（此时 PageWidth/PageHeight 已等于内容区域大小）
    page.Export svgPath

    ' 回滚（bCommit = False），还原图纸尺寸，不污染原文件
    Application.EndUndoScope scopeID, False
    Exit Sub

ErrHandler:
    ' 出错时同样回滚，然后降级为直接导出原始页面
    Application.EndUndoScope scopeID, False
    page.Export svgPath
End Sub

' 移除文件名中的非法字符
Private Function SanitizeFileName(name As String) As String
    Dim illegal As String
    Dim i As Integer
    Dim result As String

    ' Windows 文件名非法字符
    illegal = "\/:*?""<>|"
    result = name

    For i = 1 To Len(illegal)
        result = Join(Split(result, Mid(illegal, i, 1)), "_")
    Next i

    ' 去除首尾空格及点号
    result = Trim(result)
    Do While Left(result, 1) = "." : result = Mid(result, 2) : Loop
    Do While Right(result, 1) = "." : result = Left(result, Len(result) - 1) : Loop

    If result = "" Then result = "Page"

    SanitizeFileName = result
End Function

' 若目标路径已存在文件，则追加 (n) 序号返回唯一路径
Private Function UniqueFilePath(path As String) As String
    If Dir(path) = "" Then
        UniqueFilePath = path
        Exit Function
    End If

    Dim base As String
    Dim ext As String
    Dim dotPos As Long
    dotPos = InStrRev(path, ".")
    If dotPos > 0 Then
        base = Left(path, dotPos - 1)
        ext = Mid(path, dotPos)
    Else
        base = path
        ext = ""
    End If

    Dim n As Long
    n = 1
    Dim candidate As String
    Do
        candidate = base & " (" & n & ")" & ext
        n = n + 1
    Loop While Dir(candidate) <> ""

    UniqueFilePath = candidate
End Function

' 弹出文件夹选择对话框，返回选中的路径（取消则返回空字符串）
Private Function BrowseForFolder(title As String) As String
    Dim shell As Object
    Dim folder As Object

    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, title, 0, "")

    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.Path
    Else
        BrowseForFolder = ""
    End If
End Function

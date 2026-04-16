' ExportVisioToSVG.vba
' 将 Visio 文件中每个 Page 导出为 SVG 文件，文件名以 Sheet（Page）名称命名。
'
' 核心策略：先按原始图纸尺寸导出 SVG，再通过 MSXML 修改 SVG 的 viewBox 属性，
' 使图形内容充满整张图纸。全程不修改 Visio 文档。
'
' 使用方法：
'   1. 打开 Visio，按 Alt+F11 打开 VBA 编辑器
'   2. 插入 > 模块，将此代码粘贴进去
'   3. 修改下方 DEFAULT_OUTPUT_DIR 常量（或运行时弹窗选择目录）
'   4. 运行 ExportAllPagesToSVG 宏

Option Explicit

' 默认输出目录，留空则运行时弹窗选择
Private Const DEFAULT_OUTPUT_DIR As String = ""

' ==============================================================
' 入口
' ==============================================================

Public Sub ExportAllPagesToSVG()
    Dim doc As Visio.Document
    Dim outputDir As String

    If Visio.Documents.Count = 0 Then
        MsgBox "没有打开的 Visio 文档，请先打开一个 .vsdx 文件。", vbExclamation
        Exit Sub
    End If

    Set doc = Visio.ActiveDocument

    outputDir = Trim(DEFAULT_OUTPUT_DIR)
    If outputDir = "" Then
        outputDir = BrowseForFolder("请选择 SVG 文件的输出目录")
        If outputDir = "" Then
            MsgBox "未选择输出目录，操作已取消。", vbInformation
            Exit Sub
        End If
    End If

    If Right(outputDir, 1) <> "\" Then outputDir = outputDir & "\"

    Dim exported As Long
    exported = ExportPages(doc, outputDir)

    MsgBox "导出完成！共导出 " & exported & " 个 SVG 文件。" & vbCrLf & _
           "输出目录：" & outputDir, vbInformation
End Sub

' ==============================================================
' 遍历页面
' ==============================================================

Private Function ExportPages(doc As Visio.Document, outputDir As String) As Long
    Dim page As Visio.Page
    Dim count As Long
    count = 0

    For Each page In doc.Pages
        Dim safeName As String
        safeName = SanitizeFileName(page.Name)

        Dim finalPath As String
        finalPath = UniqueFilePath(outputDir & safeName & ".svg")

        ExportPageFitContent page, finalPath
        count = count + 1
    Next page

    ExportPages = count
End Function

' ==============================================================
' 导出单页并裁剪 viewBox
' ==============================================================

' 先按原始尺寸导出 SVG（不改动 Visio 文档），
' 再计算所有形状的外接矩形，更新 SVG 的 viewBox，使内容充满图纸。
Private Sub ExportPageFitContent(page As Visio.Page, svgPath As String)
    ' 1. 正常导出
    page.Export svgPath

    ' 无形状则无需调整
    If page.Shapes.Count = 0 Then Exit Sub

    ' 2. 计算形状外接矩形（Visio 绘图坐标，单位：英寸）
    Dim minX As Double, minY As Double, maxX As Double, maxY As Double
    If Not GetShapesBBox(page, minX, minY, maxX, maxY) Then Exit Sub

    ' 3. 修改 SVG 的 viewBox
    AdjustSVGViewBox svgPath, page, minX, minY, maxX, maxY
End Sub

' ==============================================================
' 计算所有形状的外接矩形
' ==============================================================

' 遍历页面所有顶层形状，返回它们的联合外接矩形。
' 使用 Visio 绘图坐标（原点在页面左下角，Y 轴向上，单位：英寸）。
' 返回 True 表示至少找到一个有效形状。
Private Function GetShapesBBox(page As Visio.Page, _
                                ByRef minX As Double, ByRef minY As Double, _
                                ByRef maxX As Double, ByRef maxY As Double) As Boolean
    Dim shp As Visio.Shape
    Dim l As Double, b As Double, r As Double, t As Double
    Dim errNum As Long
    Dim found As Boolean: found = False

    For Each shp In page.Shapes
        On Error Resume Next
        ' 4 = visBBoxDrawingCoords，使用绘图坐标系
        shp.BoundingBox 4, l, b, r, t
        errNum = Err.Number
        On Error GoTo 0

        If errNum = 0 And r > l And t > b Then
            If Not found Then
                minX = l: minY = b: maxX = r: maxY = t
                found = True
            Else
                If l < minX Then minX = l
                If b < minY Then minY = b
                If r > maxX Then maxX = r
                If t > maxY Then maxY = t
            End If
        End If
    Next shp

    GetShapesBBox = found
End Function

' ==============================================================
' 修改 SVG 的 viewBox
' ==============================================================

' 通过 MSXML 读取已导出的 SVG，根据形状外接矩形重算 viewBox，写回文件。
' 原理：
'   Visio 导出的 SVG viewBox = "0 0 PageW PageH"（SVG 内部单位）
'   通过比例换算，将 Visio 绘图坐标（英寸）映射到 SVG 坐标。
'   注意：Visio Y 轴向上，SVG Y 轴向下，需翻转。
Private Sub AdjustSVGViewBox(svgPath As String, page As Visio.Page, _
                               minX As Double, minY As Double, _
                               maxX As Double, maxY As Double)
    ' 读取页面原始尺寸（英寸）
    Dim ps As Visio.Shape: Set ps = page.PageSheet
    Dim pageW As Double: pageW = ps.CellsU("PageWidth").ResultIU
    Dim pageH As Double: pageH = ps.CellsU("PageHeight").ResultIU
    If pageW <= 0 Or pageH <= 0 Then Exit Sub

    ' 创建 MSXML 解析器
    Dim xml As Object
    On Error Resume Next
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    On Error GoTo 0
    If xml Is Nothing Then Exit Sub

    xml.async = False
    xml.validateOnParse = False
    xml.resolveExternals = False

    If Not xml.Load(svgPath) Then Exit Sub
    If xml.parseError.errorCode <> 0 Then Exit Sub

    Dim root As Object: Set root = xml.documentElement
    If root Is Nothing Then Exit Sub

    ' 读取现有 viewBox（格式："x y w h"）
    Dim vbStr As String: vbStr = root.getAttribute("viewBox")
    If vbStr = "" Then Exit Sub

    ' 清理分隔符，确保只有单个空格
    vbStr = Replace(vbStr, ",", " ")
    Do While InStr(vbStr, "  ") > 0
        vbStr = Replace(vbStr, "  ", " ")
    Loop
    Dim parts() As String: parts = Split(Trim(vbStr), " ")
    If UBound(parts) < 3 Then Exit Sub

    Dim svgFullW As Double: svgFullW = CDbl(parts(2))   ' 全页宽度（SVG 内部单位）
    Dim svgFullH As Double: svgFullH = CDbl(parts(3))   ' 全页高度（SVG 内部单位）
    If svgFullW <= 0 Or svgFullH <= 0 Then Exit Sub

    ' 换算比例：Visio 英寸 → SVG 内部单位
    Dim scaleX As Double: scaleX = svgFullW / pageW
    Dim scaleY As Double: scaleY = svgFullH / pageH

    ' 计算新 viewBox（Y 轴翻转：SVG Y = (pageH - visioY) * scaleY）
    Dim newX As Double: newX = minX * scaleX
    Dim newY As Double: newY = (pageH - maxY) * scaleY
    Dim newW As Double: newW = (maxX - minX) * scaleX
    Dim newH As Double: newH = (maxY - minY) * scaleY
    If newW <= 0 Or newH <= 0 Then Exit Sub

    ' 写入新 viewBox（使用不受区域设置影响的小数格式）
    root.setAttribute "viewBox", _
        InvFmt(newX) & " " & InvFmt(newY) & " " & InvFmt(newW) & " " & InvFmt(newH)

    xml.Save svgPath
End Sub

' 将浮点数格式化为始终使用 "." 小数点的字符串（避免区域设置影响 SVG 解析）
Private Function InvFmt(d As Double) As String
    InvFmt = Replace(Format(d, "0.######"), ",", ".")
End Function

' ==============================================================
' 工具函数
' ==============================================================

' 移除文件名中的非法字符
Private Function SanitizeFileName(name As String) As String
    Dim illegal As String
    Dim i As Integer
    Dim result As String

    illegal = "\/:*?""<>|"
    result = name

    For i = 1 To Len(illegal)
        result = Join(Split(result, Mid(illegal, i, 1)), "_")
    Next i

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

    Dim n As Long: n = 1
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

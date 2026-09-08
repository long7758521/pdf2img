Attribute VB_Name = "PDF转图片插入Word"
' =============================================================================
'  PDF转图片插入Word 示例宏
'  配合 pdf2img.exe 使用：获取页数 -> 转换 -> 读取退出码 -> 逐张插入图片 -> 清理
'  使用前请修改“可配置常量”区域，将路径改为你的实际环境。
'  导入方式：Word 按 Alt+F11 -> 菜单“文件”-> 导入文件 -> 选中本 .bas
' =============================================================================
Option Explicit

' ---------------------------------------------------------------------------
' 可配置常量
' ---------------------------------------------------------------------------
Private Const PDF2IMG_EXE As String = "C:\Tools\pdf2img.exe"   ' pdf2img.exe 绝对路径
Private Const PDF_PATH As String = "D:\标书\输入\tender.pdf"    ' 待转换 PDF 绝对路径
Private Const OUT_DIR As String = "D:\标书\临时图片\"           ' 图片输出目录（末尾带 \
Private Const DPI As Long = 150
Private Const IMG_EXT As String = "png"

' ---------------------------------------------------------------------------
' 主流程宏：调用本宏即可完成【转换 + 插入Word光标处 + 清理】
' ---------------------------------------------------------------------------
Public Sub PDF转图片插入Word()

    Dim shell As Object
    Dim n As Long        ' PDF 总页数
    Dim exitCode As Long ' 工具退出码
    Dim i As Long

    ' 0. 校验文件存在
    If Dir(PDF2IMG_EXE) = "" Then
        MsgBox "找不到 pdf2img.exe: " & PDF2IMG_EXE, vbCritical
        Exit Sub
    End If
    If Dir(PDF_PATH) = "" Then
        MsgBox "找不到 PDF 文件: " & PDF_PATH, vbCritical
        Exit Sub
    End If

    Set shell = CreateObject("WScript.Shell")

    ' 1. 获取 PDF 页数（精确循环控制）
    n = GetPdfPageCount(shell)
    If n <= 0 Then
        MsgBox "获取页数失败，请检查 PDF 文件。", vbCritical
        Exit Sub
    End If
    MsgBox "PDF 共 " & n & " 页，即将开始转换并插入 Word。", vbInformation

    ' 2. 执行转换，等待结束并读取退出码
    exitCode = ConvertPdf(shell)
    If exitCode <> 0 Then
        MsgBox "转换失败，退出码 " & exitCode & "（" & ExitCodeMsg(exitCode) & "）。", vbCritical
        Exit Sub
    End If

    ' 3. 逐张插入图片到光标处
    For i = 1 To n
        Dim imgPath As String
        imgPath = OUT_DIR & BaseName(PDF_PATH) & "_" & i & "." & IMG_EXT
        Selection.InlineShapes.AddPicture FileName:=imgPath
        Selection.TypeParagraph
    Next i

    ' 4. 清理临时图片（Kill 删除）
    For i = 1 To n
        On Error Resume Next
        Kill OUT_DIR & BaseName(PDF_PATH) & "_" & i & "." & IMG_EXT
        On Error GoTo 0
    Next i

    MsgBox "完成！已插入 " & n & " 张图片并清理缓存。", vbInformation

End Sub

' ---------------------------------------------------------------------------
' 获取 PDF 总页数（读取 stdout 的纯数字）
' ---------------------------------------------------------------------------
Private Function GetPdfPageCount(shell As Object) As Long
    Dim exec As Object
    Dim cmdLine As String

    cmdLine = """" & PDF2IMG_EXE & """ """ & PDF_PATH & """ --get-page-count --quiet"
    Set exec = shell.Exec(cmdLine)
    GetPdfPageCount = Val(exec.StdOut.ReadLine)
End Function

' ---------------------------------------------------------------------------
' 执行完整转换，返回进程退出码
' ---------------------------------------------------------------------------
Private Function ConvertPdf(shell As Object) As Long
    Dim cmdLine As String

    cmdLine = """" & PDF2IMG_EXE & """ """ & PDF_PATH & _
              """ --output """ & OUT_DIR & _
              """ --dpi " & DPI & _
              " --format " & IMG_EXT & " --quiet"
    ConvertPdf = shell.Run(cmdLine, 0, True)
End Function

' ---------------------------------------------------------------------------
' 退出码 -> 含义
' ---------------------------------------------------------------------------
Private Function ExitCodeMsg(code As Long) As String
    Select Case code
        Case 1: ExitCodeMsg = "输入参数错误"
        Case 2: ExitCodeMsg = "PDF 文件不存在或无法访问"
        Case 3: ExitCodeMsg = "PDF 损坏、加密或无法解析"
        Case 4: ExitCodeMsg = "某页渲染失败"
        Case 5: ExitCodeMsg = "输出目录创建失败或磁盘空间不足"
        Case 6: ExitCodeMsg = "写入图片失败"
        Case 7: ExitCodeMsg = "未预期的运行时错误"
        Case Else: ExitCodeMsg = "未知错误"
    End Select
End Function

' ---------------------------------------------------------------------------
' 取文件名主干（不含扩展名）
' ---------------------------------------------------------------------------
Private Function BaseName(fullPath As String) As String
    BaseName = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
    If InStr(BaseName, ".") > 0 Then
        BaseName = Mid$(BaseName, 1, InStr(BaseName, ".") - 1)
    End If
End Function
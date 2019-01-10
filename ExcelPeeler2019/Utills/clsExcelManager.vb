Imports System
Imports Autodesk
Imports Autodesk.Revit
Imports XLS = Microsoft.Office.Interop.Excel
Imports System.Runtime

Public Class ClsExcelManager

    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function

    ''' <summary>
    ''' エクセルオブジェクトとシートを取得する
    ''' </summary>
    ''' <param name="XLSapp">エクセルアプリケーション</param>
    ''' <param name="sh">現在アクティブなシート</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetExcelApp(ByRef XLSapp As XLS.Application, ByRef sh As XLS.Worksheet) As Integer


        '隠れたプロセスは削除する

        Dim wb As XLS.Workbook = Nothing
        Try
            Dim bFindXLS As Boolean = False

            Do While bFindXLS = False
                XLSapp = CType(GetObject(, "Excel.Application"), XLS.Application)
                If XLSapp.Visible = False Then
                    XLSapp.Quit()
                    Dim xlHWND As Integer = XLSapp.Hwnd
                    Dim ProcIdXL As Integer = 0
                    'get the process ID
                    GetWindowThreadProcessId(xlHWND, ProcIdXL)
                    Dim ps As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessById(ProcIdXL)
                    ps.Kill()
                    InteropServices.Marshal.ReleaseComObject(XLSapp)
                    XLSapp = Nothing
                Else
                    bFindXLS = True
                End If
            Loop

            'XLSapp.Visible = True
            If XLSapp.Workbooks.Count = 0 Then
                wb = CType(XLSapp.Workbooks.Add, XLS.Workbook)
            Else
                wb = CType(XLSapp.ActiveWorkbook, XLS.Workbook)
            End If

            '現在のシート
            If wb.Worksheets.Count = 0 Then
                sh = CType(wb.Worksheets.Add, XLS.Worksheet)
            Else
                sh = CType(wb.ActiveSheet, XLS.Worksheet)
            End If


        Catch ex As Exception

            ReleaseComObject(XLSapp)
            ReleaseComObject(wb)
            ReleaseComObject(sh)

            XLSapp = Nothing
            wb = Nothing
            sh = Nothing
            Return -1

            ''起動していない場合は作成する
            'XlsApp = CreateObject("Excel.Application")
            'XlsApp.Visible = True
            ''シートを新しく作成する
            'wb = XlsApp.Workbooks.Add
        End Try

        'wb = Nothing
        ReleaseComObject(wb)
        If sh Is Nothing Then
            ReleaseComObject(XLSapp)
            XLSapp = Nothing
            Return -2
            Exit Function
        End If

        Return 1

    End Function

    Public Function GetExcelApp(ByRef XLSapp As XLS.Application, ByRef wb As XLS.Workbook) As Integer

        Try

            Dim bFindXLS As Boolean = False

            Do While bFindXLS = False
                XLSapp = CType(GetObject(, "Excel.Application"), XLS.Application)
                If XLSapp.Visible = False Then
                    XLSapp.Quit()
                    Dim xlHWND As Integer = XLSapp.Hwnd
                    Dim ProcIdXL As Integer = 0
                    'get the process ID
                    GetWindowThreadProcessId(xlHWND, ProcIdXL)
                    Dim ps As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessById(ProcIdXL)
                    ps.Kill()
                    InteropServices.Marshal.ReleaseComObject(XLSapp)
                    XLSapp = Nothing
                Else
                    bFindXLS = True
                End If
            Loop



            'XLSapp = CType(GetObject(, "Excel.Application"), XLS.Application)

            ''ここでエクセルのプロセスが複数ある場合がある。
            'If XLSapp.Visible = False Then
            '    'XLSapp.Visible = True
            '    XLSapp.Quit()
            '    MsgBox("エクセルを特定できません。プロセスが複数存在するかもしれません。" + vbCrLf + _
            '           "エクセルをいったん終了したうえで、タスクマネージャーを起動し" + vbCrLf + _
            '           "EXCEL.exeのプロセスをすべて終了してから再度起動してください。")
            '    InteropServices.Marshal.ReleaseComObject(XLSapp)

            '    'XLSapp = Nothing
            '    Return -1

            'End If


            'XLSapp.Visible = True
            If XLSapp.Workbooks.Count = 0 Then
                wb = CType(XLSapp.Workbooks.Add, XLS.Workbook)
            Else
                wb = CType(XLSapp.ActiveWorkbook, XLS.Workbook)
            End If



        Catch ex As Exception

            ReleaseComObject(XLSapp)
            ReleaseComObject(wb)

            XLSapp = Nothing
            wb = Nothing
            Return -1

        End Try

        If wb Is Nothing Then
            ReleaseComObject(XLSapp)
            XLSapp = Nothing
            Return -2
            Exit Function
        End If

        Return 1

    End Function


    Private Sub testExcelRunning()
        On Error Resume Next
        ' GetObject called without the first argument returns a
        ' reference to an instance of the application. If the
        ' application is not already running, an error occurs.
        Dim excelObj As Object = GetObject(, "Excel.Application")
        If Err.Number = 0 Then
            MsgBox("Excel is running")
        Else
            MsgBox("Excel is not running")
        End If
        Err.Clear()
        excelObj = Nothing
    End Sub

    Public Function GetExcelApp2(ByRef XLSapp As XLS.Application) As Integer
        On Error Resume Next

        Dim excelObj As Object = GetObject(, "Excel,Application")
        Dim ret As Integer = 0
        If Err.Number = 0 Then
            MsgBox("Excel is running")
            XLSapp = excelObj.Application
            ret = 1
        Else
            MsgBox("Excel is not running")
            ret = 0
        End If

        Err.Clear()

        Return ret


    End Function


    ''' <summary>
    ''' 現在選択されている行番号のリストを返す
    ''' </summary>
    ''' <param name="xApp"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSelectedRowList(ByVal xApp As XLS.Application) As List(Of Integer)

        Dim rngSel As XLS.Range = CType(xApp.Selection, XLS.Range)
        Dim areaSel As XLS.Areas = rngSel.Areas
        Dim lstSelNumber As New List(Of Integer)
        For Each rng1 As XLS.Range In areaSel
            For Each rw As XLS.Range In rng1.Rows
                Dim selRow As Integer = rw.Row
                If lstSelNumber.Contains(selRow) = False Then
                    lstSelNumber.Add(selRow)
                End If
            Next
        Next
        ReleaseComObject(rngSel)
        ReleaseComObject(areaSel)

        lstSelNumber.Sort()
        Return lstSelNumber


    End Function

    ''' <summary>
    ''' 現在選択されている選択されたセルの整数の値を返す
    ''' </summary>
    ''' <param name="xApp"></param>
    ''' <returns></returns>
    Public Function GetSelectedIntVals(ByVal xApp As XLS.Application) As List(Of Integer)
        Dim rngSel As XLS.Range = CType(xApp.Selection, XLS.Range)
        Dim areaSel As XLS.Areas = rngSel.Areas
        Dim lstSelNumber As New List(Of Integer)
        For Each rng1 As XLS.Range In areaSel
            For Each cel1 As XLS.Range In rng1.Cells
                Dim val1 As Object = cel1.Value
                Try
                    Dim intVal1 As Integer = Integer.Parse(val1)
                    If IsNothing(intVal1) = True Then
                        Continue For
                    End If
                    If lstSelNumber.Contains(intVal1) = False Then
                        lstSelNumber.Add(intVal1)
                    End If

                Catch ex As Exception
                    Continue For
                End Try
            Next
        Next
        ReleaseComObject(rngSel)
        ReleaseComObject(areaSel)

        lstSelNumber.Sort()
        Return lstSelNumber

    End Function

    ''' <summary>
    ''' COMオブジェクトの解放
    ''' </summary>
    ''' <param name="objCOM"></param>
    Public Sub ReleaseComObject(ByRef objCOM As Object)
        Try
            ' ランタイム呼び出し可能ラッパーの参照カウントをデクリメント
            If ((Not objCOM Is Nothing) AndAlso
                (System.Runtime.InteropServices.Marshal.IsComObject(objCOM))) Then
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objCOM)
            End If
        Finally
            ' 参照を解除する
            objCOM = Nothing
        End Try

    End Sub


    ''' <summary>
    ''' セルのフォーマット
    ''' </summary>
    ''' <param name="xlRange"></param>
    Public Sub FormatCells1(ByVal xlRange As XLS.Range)

        xlRange.Borders(XLS.XlBordersIndex.xlDiagonalDown).LineStyle = XLS.Constants.xlNone
        xlRange.Borders(XLS.XlBordersIndex.xlDiagonalUp).LineStyle = XLS.Constants.xlNone
        With xlRange.Borders(XLS.XlBordersIndex.xlEdgeLeft)
            .LineStyle = XLS.XlLineStyle.xlContinuous
            .ColorIndex = XLS.Constants.xlAutomatic
            .TintAndShade = 0
            .Weight = XLS.XlBorderWeight.xlThick
        End With
        With xlRange.Borders(XLS.XlBordersIndex.xlEdgeTop)
            .LineStyle = XLS.XlLineStyle.xlContinuous
            .ColorIndex = XLS.Constants.xlAutomatic
            .TintAndShade = 0
            .Weight = XLS.XlBorderWeight.xlThick
        End With
        With xlRange.Borders(XLS.XlBordersIndex.xlEdgeBottom)
            .LineStyle = XLS.XlLineStyle.xlContinuous
            .ColorIndex = XLS.Constants.xlAutomatic
            .TintAndShade = 0
            .Weight = XLS.XlBorderWeight.xlThick
        End With
        With xlRange.Borders(XLS.XlBordersIndex.xlEdgeRight)
            .LineStyle = XLS.XlLineStyle.xlContinuous
            .ColorIndex = XLS.Constants.xlAutomatic
            .TintAndShade = 0
            .Weight = XLS.XlBorderWeight.xlThick
        End With
        With xlRange.Borders(XLS.XlBordersIndex.xlInsideVertical)
            .LineStyle = XLS.XlLineStyle.xlContinuous
            .ColorIndex = XLS.Constants.xlAutomatic
            .TintAndShade = 0
            .Weight = XLS.XlBorderWeight.xlHairline
        End With
        With xlRange.Borders(XLS.XlBordersIndex.xlInsideHorizontal)
            .LineStyle = XLS.XlLineStyle.xlContinuous
            .ColorIndex = XLS.Constants.xlAutomatic
            .TintAndShade = 0
            .Weight = XLS.XlBorderWeight.xlHairline
        End With

        ReleaseComObject(xlRange)


    End Sub


    ''' <summary>
    ''' ダブらないワークシートの名前を返す
    ''' </summary>
    ''' <param name="ws"></param>
    ''' <param name="PrefixName"></param>
    ''' <returns></returns>
    Public Function GetSafeWsName(ByVal ws As XLS.Worksheet, ByVal PrefixName As String) As String

        Dim wb As XLS.Workbook = ws.Application.ActiveWorkbook
        Dim lstSheetNames As New List(Of String)
        For Each ws1 As XLS.Worksheet In wb.Worksheets
            lstSheetNames.Add(ws1.Name.ToUpper)
        Next

        Dim i As Integer = 0
        Dim bExist As Boolean = False
        Dim newName As String = PrefixName
        Do While bExist = False
            If i <> 0 Then
                newName = newName + i.ToString
            End If
            If lstSheetNames.Contains(newName.ToUpper) = False Then
                bExist = True
                Exit Do
            End If
            i += 1
        Loop

        Return newName

    End Function


End Class

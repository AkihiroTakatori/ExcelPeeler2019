Imports System
Imports Autodesk
Imports Autodesk.Revit
Imports XLS = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop


Public Class ClsExcelUtils


    ''' <summary>
    ''' エクセルに出力
    ''' </summary>
    ''' <param name="lstParams">パラメーターのリスト</param>
    ''' <param name="lstElements">要素のリスト</param>
    ''' <param name="sh">エクセルのシート</param>
    ''' <param name="dbDoc">ドキュメント</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExportToExcel(ByVal lstParams As List(Of String), ByVal lstElements As List(Of DB.Element), ByVal sh As XLS.Worksheet, ByVal dbDoc As DB.Document) As Integer

        Dim iRow As Integer
        Dim iCol As Integer
        'タイトル行の作成
        '(1,1)="UID"
        sh.Cells(1, 1).Value = "UID"
        '(1,2)から後ろ
        For iCol = 2 To lstParams.Count + 1
            sh.Cells(1, iCol).Value = lstParams.Item(iCol - 2)
        Next

        '2行目以降はパラメータの値

        iRow = 2
        For Each elm As DB.Element In lstElements
            iCol = 1
            '(iRow,1)はUID
            sh.Cells(iRow, iCol).Value = elm.UniqueId.ToString
            '(iRow,iCol)はパラメータにしたがって
            For i As Integer = 0 To lstParams.Count - 1
                iCol += 1
                Dim strPname As String = lstParams.Item(i)
                'パラメータ情報の取得
                Dim strValue As String = ""
                If strPname = "ファミリ" Then
                    If TypeOf elm Is DB.Wall Then
                        Dim wElm As DB.Wall = TryCast(elm, DB.Wall)
                        If wElm.WallType.Kind = DB.WallKind.Basic Then
                            strValue = "標準壁"
                        ElseIf wElm.WallType.Kind = DB.WallKind.Curtain Then
                            strValue = "カーテンウォール"
                        ElseIf wElm.WallType.Kind = DB.WallKind.Stacked Then
                            strValue = "スタックウォール"
                        Else
                            strValue = "不明"
                        End If

                    ElseIf TypeOf elm Is DB.WallType Then
                        Dim wTp As DB.WallType = TryCast(elm, DB.WallType)
                        If wTp.Kind = DB.WallKind.Basic Then
                            strValue = "標準壁"
                        ElseIf wTp.Kind = DB.WallKind.Curtain Then
                            strValue = "カーテンウォール"
                        ElseIf wTp.Kind = DB.WallKind.Stacked Then
                            strValue = "スタックウォール"
                        Else
                            strValue = "不明"
                        End If
                        '================床==================================
                    ElseIf TypeOf elm Is DB.FloorType Then
                        Dim fTp As DB.FloorType = TryCast(elm, DB.FloorType)
                        If fTp.IsFoundationSlab = True Then
                            strValue = "床"
                        Else
                            strValue = "基礎スラブ"
                        End If
                    ElseIf TypeOf elm Is DB.Floor Then
                        Dim fElm As DB.Floor = TryCast(elm, DB.Floor)
                        If fElm.FloorType.IsFoundationSlab = True Then
                            strValue = "床"
                        Else
                            strValue = "基礎スラブ"
                        End If
                        '================ファミリ（構造系と意匠柱)==================================
                    ElseIf TypeOf elm Is DB.FamilySymbol Then
                        Dim fs As DB.FamilySymbol = TryCast(elm, DB.FamilySymbol)
                        strValue = fs.Family.Name

                    ElseIf TypeOf elm Is DB.FamilyInstance Then
                        Dim fi As DB.FamilyInstance = TryCast(elm, DB.FamilyInstance)
                        strValue = fi.Symbol.Family.Name





                    End If

                ElseIf strPname = "タイプ" Then
                    '================壁=================================
                    If TypeOf elm Is DB.Wall Then

                        Dim wElm As DB.Wall = TryCast(elm, DB.Wall)
                        strValue = wElm.WallType.Name

                    ElseIf TypeOf elm Is DB.WallType Then
                        Dim wTp As DB.WallType = TryCast(elm, DB.WallType)
                        strValue = wTp.Name
                        '================床==================================
                    ElseIf TypeOf elm Is DB.FloorType Then
                        Dim fTp As DB.FloorType = TryCast(elm, DB.FloorType)
                        strValue = fTp.Name
                    ElseIf TypeOf elm Is DB.Floor Then
                        Dim fElm As DB.Floor = TryCast(elm, DB.Floor)
                        strValue = fElm.FloorType.Name
                        '================ファミリ（構造系と意匠柱)==================================
                    ElseIf TypeOf elm Is DB.FamilySymbol Then
                        Dim fs As DB.FamilySymbol = TryCast(elm, DB.FamilySymbol)
                        strValue = fs.Name

                    ElseIf TypeOf elm Is DB.FamilyInstance Then
                        Dim fi As DB.FamilyInstance = TryCast(elm, DB.FamilyInstance)
                        strValue = fi.Symbol.Name
                    End If

                Else
                    'パラメータの取得
                    Dim prm As DB.Parameter = elm.LookupParameter(strPname)
                    If prm Is Nothing Then
                        Continue For
                    End If
                    'パラメータタイプにしたがって出力する文字列は異なる
                    Select Case prm.StorageType
                        Case DB.StorageType.Double
                            strValue = prm.AsValueString

                        Case DB.StorageType.ElementId
                            Dim idTemp As DB.ElementId = prm.AsElementId
                            Try
                                'Dim elmTemp As DB.Element = dbDoc.Element(idTemp)
                                Dim elmTemp As DB.Element = dbDoc.GetElement(idTemp)
                                strValue = elmTemp.Name
                            Catch ex As Exception

                            End Try

                        Case DB.StorageType.Integer
                            strValue = prm.AsValueString
                            If strValue = "" Then
                                strValue = prm.AsInteger.ToString
                            End If


                        Case DB.StorageType.String
                            strValue = prm.AsString

                        Case Else

                    End Select
                End If
                If strValue Is Nothing Then
                    Continue For
                End If

                'シートに記入
                strValue = strValue.Trim
                If strValue = "" Then
                    Continue For
                End If
                sh.Cells(iRow, iCol).Value = strValue


            Next
            iRow += 1
        Next

        Return 1

    End Function


    ''' <summary>
    ''' エクセルを取得（開いてない場合は新たに開く,シートは新たに追加する)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetExcelWorkBook(Optional isAlwaysNewBook As Boolean = True) As Excel.Workbook

        '******************************************
        'エクセルを取得
        '現在起動しているExcelオブジェクトを取得する
        '******************************************

        Dim XlsApp As Excel.Application
        Dim wb As XLS.Workbook
        Try
            XlsApp = GetObject(, "Excel.Application")
            XlsApp.Visible = True
        Catch ex As Exception

            '起動していない場合は作成する
            XlsApp = CreateObject("Excel.Application")
            XlsApp.Visible = True
        End Try
        'シートを新しく作成する
        If isAlwaysNewBook = True Then
            wb = XlsApp.Workbooks.Add
        Else
            '今のwbを取得
            wb = XlsApp.ActiveWorkbook
            If wb Is Nothing Then
                wb = XlsApp.Workbooks.Add
            End If
        End If
        XlsApp = Nothing

        '現在のシート
        If wb Is Nothing Then
            Return Nothing
        End If
        Return wb

    End Function

    ''' <summary>
    ''' 現在起動しているワークブックを取得
    ''' ないばあいはNothing
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCurrentExcelWorkBook(Optional bActiveExcel As Boolean = True) As Excel.Workbook

        '******************************************
        'エクセルを取得
        '現在起動しているExcelオブジェクトを取得する
        '******************************************

        Dim XlsApp As Excel.Application
        Dim wb As XLS.Workbook
        Try
            XlsApp = GetObject(, "Excel.Application")
            If bActiveExcel = True Then
                XlsApp.Visible = True

            End If
        Catch ex As Exception

            '起動していない場合は終了する
            Return Nothing
        End Try

        wb = XlsApp.ActiveWorkbook
        XlsApp = Nothing

        If wb Is Nothing Then
            Return Nothing
        End If
        Return wb

    End Function


    ''' <summary>
    ''' ワークシートの1行目にタイトルを書き込む
    ''' </summary>
    ''' <param name="sh"></param>
    ''' <param name="ParameterList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExportParamHeader(ByVal sh As XLS.Worksheet, ByVal ParameterList As List(Of String)) As Integer

        Dim iCol As Integer = 1
        sh.Rows.RowHeight = 35.0# '一旦すべての行の高さを35にする

        'UID,FAMILY,TYPE,Param1,Param2,,,,,,
        FormatTitleCell(1, 1, "UID", 10, sh)
        FormatTitleCell(1, 2, "Family" + vbCrLf + "Category", 10, sh)
        FormatTitleCell(1, 3, "Type", 10, sh)
        iCol = 4
        For index As Integer = 0 To ParameterList.Count - 1
            Dim prmTemp As String = ParameterList.Item(index)
            '書式の変更
            FormatParamTitleCell(1, iCol, prmTemp, 10, sh)
            '次
            iCol += 1
        Next

        Return iCol - 1

    End Function

    Public Function ExportElementIT(ByVal sh As XLS.Worksheet,
                                    ByVal ElementList As List(Of DB.Element),
                                    ByVal ParameterNames As List(Of String)) As Integer

        Dim dbDoc As DB.Document = ElementList.Item(0).Document
        Dim iRow As Integer = 2
        For Each elmTemp As DB.Element In ElementList
            'sh.Rows(iRow).Select()
            'このelmTempはインスタンスであることが前提なのでタイプを取得する

            Try
                'タイプの取得
                Dim typTemp As DB.Element = dbDoc.GetElement(elmTemp.GetTypeId)

                '(1)UID
                sh.Cells(iRow, 1).Value = elmTemp.UniqueId
                '(2) ファミリ名/カテゴリ名
                If TypeOf elmTemp Is DB.FamilyInstance Then
                    Dim faminst As DB.FamilyInstance = DirectCast(elmTemp, DB.FamilyInstance)
                    Dim famName As String = faminst.Symbol.Family.Name
                    If String.IsNullOrEmpty(famName) = False Then
                        sh.Cells(iRow, 2).Value = famName
                    End If
                ElseIf TypeOf elmTemp Is DB.FamilySymbol Then
                    'ここは使わないはず
                    Dim famsym As DB.FamilySymbol = DirectCast(elmTemp, DB.FamilySymbol)
                    Dim famName As String = famsym.Family.Name
                    If String.IsNullOrEmpty(famName) = False Then
                        sh.Cells(iRow, 2).Value = famName
                    End If
                ElseIf TypeOf elmTemp Is DB.View Then
                    Dim viewTemp As DB.View = DirectCast(elmTemp, DB.View)
                    Dim vfType As DB.ViewFamilyType = dbDoc.GetElement(viewTemp.GetTypeId)
                    sh.Cells(iRow, 2).Value = vfType.Name
                Else
                    'その他の場合はカテゴリの名前を入れておく
                    sh.Cells(iRow, 2).Value = elmTemp.Category.Name
                End If

                '(3)名前
                Dim elmName As String = elmTemp.Name
                If String.IsNullOrEmpty(elmName) = False Then
                    sh.Cells(iRow, 3).Value = elmName
                End If


                '(4)以降はParameterNamesに従う
                For j As Integer = 0 To ParameterNames.Count - 1
                    Dim sParamName As String = ParameterNames(j)
                    Dim iCol As Integer = 4 + j
                    Dim prmTemp As DB.Parameter = Nothing
                    If sParamName.StartsWith("T:") = True And IsNothing(typTemp) = False Then
                        sParamName = sParamName.Remove(0, 2)
                        If sParamName = "タイプGUID" Then
                            Dim sV As String = typTemp.UniqueId
                            sh.Cells(iRow, iCol).Value = sV
                        Else
                            prmTemp = typTemp.LookupParameter(sParamName)
                        End If
                    Else
                        'インスタンスパラメータ----ToRoomなどの場合はprmTemp=nothingのままなので、最後の部分は処理されない。
                        If sParamName = "ToRoom" Then
                            Dim fi As DB.FamilyInstance = TryCast(elmTemp, DB.FamilyInstance)
                            If fi IsNot Nothing Then
                                Dim rm As DB.Architecture.Room = fi.ToRoom
                                If rm IsNot Nothing Then
                                    Dim sV As String = rm.UniqueId
                                    sh.Cells(iRow, iCol).Value = sV
                                End If
                            End If
                        ElseIf sParamName = "FromRoom" Then
                            Dim fi As DB.FamilyInstance = TryCast(elmTemp, DB.FamilyInstance)
                            If fi IsNot Nothing Then
                                Dim rm As DB.Architecture.Room = fi.FromRoom
                                If rm IsNot Nothing Then
                                    Dim sV As String = rm.UniqueId
                                    sh.Cells(iRow, iCol).Value = sV
                                    'sh.Cells(iRow, iCol).Select()
                                End If
                            End If
                        ElseIf sParamName = "Room" Then
                            Dim fi As DB.FamilyInstance = TryCast(elmTemp, DB.FamilyInstance)
                            If fi IsNot Nothing Then
                                Dim rm As DB.Architecture.Room = fi.Room
                                If rm IsNot Nothing Then
                                    Dim sV As String = rm.UniqueId
                                    sh.Cells(iRow, iCol).Value = sV
                                    'sh.Cells(iRow, iCol).Select()
                                End If
                            End If
                        ElseIf sParamName = "Space" Then
                            Dim fi As DB.FamilyInstance = TryCast(elmTemp, DB.FamilyInstance)
                            If fi IsNot Nothing Then
                                Dim spa As Autodesk.Revit.DB.Mechanical.Space = fi.Space
                                If spa IsNot Nothing Then
                                    Dim sV As String = spa.UniqueId
                                    sh.Cells(iRow, iCol).Value = sV
                                    'sh.Cells(iRow, iCol).Select()
                                End If
                            End If
                        ElseIf sParamName = "Host" Then
                            Dim fi As DB.FamilyInstance = TryCast(elmTemp, DB.FamilyInstance)
                            If fi IsNot Nothing Then
                                Dim elmHost As DB.Element = fi.Host
                                If elmHost IsNot Nothing Then
                                    Dim hV As String = elmHost.UniqueId
                                    sh.Cells(iRow, iCol).Value = hV
                                    'sh.Cells(iRow, iCol).Select()
                                End If
                            End If
                        ElseIf sParamName = "SpaceName" Then
                            Dim fi As DB.FamilyInstance = TryCast(elmTemp, DB.FamilyInstance)
                            If fi IsNot Nothing Then
                                Dim spa As Autodesk.Revit.DB.Mechanical.Space = fi.Space
                                If spa IsNot Nothing Then
                                    Dim prmSp As DB.Parameter = spa.Parameter(DB.BuiltInParameter.ROOM_NAME)
                                    Dim sV As String = ""
                                    If IsNothing(prmSp) = False Then
                                        sV = prmSp.AsString
                                    End If
                                    sh.Cells(iRow, iCol).Value = sV
                                    'sh.Cells(iRow, iCol).Select()
                                End If
                            End If
                        ElseIf sParamName = "ID" Then
                            sh.Cells(iRow, iCol).Value = elmTemp.Id.IntegerValue
                        Else
                            prmTemp = elmTemp.LookupParameter(sParamName)
                        End If
                    End If

                    '見つかった場合
                    If IsNothing(prmTemp) = False Then
                        'パラメータの値
                        Dim sParamValue As String = ""
                        Select Case prmTemp.StorageType
                            Case DB.StorageType.Double
                                Dim sVal As Double = prmTemp.AsDouble
                                sParamValue = prmTemp.AsValueString
                            Case DB.StorageType.ElementId
                                Dim valueElm As DB.Element = elmTemp.Document.GetElement(prmTemp.AsElementId)
                                If valueElm IsNot Nothing Then
                                    sParamValue = valueElm.Name
                                End If

                            Case DB.StorageType.Integer
                                Dim sVal As Integer = prmTemp.AsInteger
                                sParamValue = sVal.ToString

                            Case DB.StorageType.String
                                sParamValue = prmTemp.AsString
                                sh.Cells(iRow, iCol).NumberFormatLocal = "@"

                        End Select
                        If String.IsNullOrEmpty(sParamValue) = False Then
                            sh.Cells(iRow, iCol).Value = sParamValue
                            'sh.Cells(iRow, iCol).Select()
                        End If
                    End If
                Next
            Catch ex As Exception
                Continue For
            End Try
            iRow += 1

        Next

        sh.Cells(2, 1).select()
        Return iRow - 1

    End Function


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sh"></param>
    ''' <param name="ElementList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExportElement(ByVal sh As XLS.Worksheet, ByVal ElementList As List(Of DB.Element)) As Integer

        Dim dbDoc As DB.Document = ElementList(0).Document
        Dim iRow As Integer = 2
        For Each elmTemp As DB.Element In ElementList
            'sh.Cells(iRow, 1).Select()
            'sh.Rows(iRow).Select()
            Try
                sh.Cells(iRow, 1).value = elmTemp.UniqueId
                '要素の名前
                Dim elmName As String = elmTemp.Name
                If String.IsNullOrEmpty(elmName) = False Then
                    sh.Cells(iRow, 3).Value = elmName
                End If
                'ファミリ名を取得
                If TypeOf elmTemp Is DB.FamilyInstance Then
                    Dim faminst As DB.FamilyInstance = DirectCast(elmTemp, DB.FamilyInstance)
                    Dim famName As String = faminst.Symbol.Family.Name
                    If String.IsNullOrEmpty(famName) = False Then
                        sh.Cells(iRow, 2).Value = famName
                    End If
                ElseIf TypeOf elmTemp Is DB.FamilySymbol Then
                    Dim famsym As DB.FamilySymbol = DirectCast(elmTemp, DB.FamilySymbol)
                    Dim famName As String = famsym.Family.Name
                    If String.IsNullOrEmpty(famName) = False Then
                        sh.Cells(iRow, 2).Value = famName
                    End If
                ElseIf TypeOf elmTemp Is DB.View Then
                    Dim viewTemp As DB.View = DirectCast(elmTemp, DB.View)
                    Dim vfType As DB.ViewFamilyType = dbDoc.GetElement(viewTemp.GetTypeId)
                    sh.Cells(iRow, 2).Value = vfType.Name
                Else
                    'その他の場合はカテゴリの名前を入れておく
                    sh.Cells(iRow, 2).Value = elmTemp.Category.Name
                End If

                '4列目以降は1行目のセルの値をパラメータの名前として読み込んで値を取得する
                'ToRoom FromRoom Room Spaceはキーワードであり、プロパティから取得する
                Dim iCol As Integer = 4
                Dim sParamName As String = sh.Cells(1, iCol).Value
                Do While String.IsNullOrEmpty(sParamName) = False

                    If sParamName = "ToRoom" Then
                        Dim fi As DB.FamilyInstance = TryCast(elmTemp, DB.FamilyInstance)
                        If fi IsNot Nothing Then
                            Dim rm As DB.Architecture.Room = fi.ToRoom
                            If rm IsNot Nothing Then
                                Dim sV As String = rm.UniqueId
                                sh.Cells(iRow, iCol).Value = sV
                                'sh.Cells(iRow, iCol).Select()
                            End If
                        End If
                    ElseIf sParamName = "FromRoom" Then
                        Dim fi As DB.FamilyInstance = TryCast(elmTemp, DB.FamilyInstance)
                        If fi IsNot Nothing Then
                            Dim rm As DB.Architecture.Room = fi.FromRoom
                            If rm IsNot Nothing Then
                                Dim sV As String = rm.UniqueId
                                sh.Cells(iRow, iCol).Value = sV
                                'sh.Cells(iRow, iCol).Select()
                            End If
                        End If
                    ElseIf sParamName = "Room" Then
                        Dim fi As DB.FamilyInstance = TryCast(elmTemp, DB.FamilyInstance)
                        If fi IsNot Nothing Then
                            Dim rm As DB.Architecture.Room = fi.Room
                            If rm IsNot Nothing Then
                                Dim sV As String = rm.UniqueId
                                sh.Cells(iRow, iCol).Value = sV
                                'sh.Cells(iRow, iCol).Select()
                            End If
                        End If
                    ElseIf sParamName = "Space" Then
                        Dim fi As DB.FamilyInstance = TryCast(elmTemp, DB.FamilyInstance)
                        If fi IsNot Nothing Then
                            Dim spa As Autodesk.Revit.DB.Mechanical.Space = fi.Space()
                            If spa IsNot Nothing Then
                                Dim sV As String = spa.UniqueId
                                sh.Cells(iRow, iCol).Value = sV
                                'sh.Cells(iRow, iCol).Select()
                            End If
                        End If
                    ElseIf sParamName = "Host" Then
                        Dim fi As DB.FamilyInstance = TryCast(elmTemp, DB.FamilyInstance)
                        If fi IsNot Nothing Then
                            Dim elmHost As DB.Element = fi.Host
                            If elmHost IsNot Nothing Then
                                Dim hV As String = elmHost.UniqueId
                                sh.Cells(iRow, iCol).Value = hV
                                'sh.Cells(iRow, iCol).Select()
                            End If
                        End If
                    ElseIf sParamName = "ID" Then
                        sh.Cells(iRow, iCol).Value = elmTemp.Id.IntegerValue
                    Else
                        'タイプパラメーターの処理
                        Dim prmTemp As DB.Parameter = elmTemp.LookupParameter(sParamName)
                        If prmTemp IsNot Nothing Then
                            'パラメータの値
                            Dim sParamValue As String = ""
                            Select Case prmTemp.StorageType
                                Case DB.StorageType.Double
                                    Dim sVal As Double = prmTemp.AsDouble
                                    sParamValue = prmTemp.AsValueString
                                Case DB.StorageType.ElementId
                                    'Dim valueElm As DB.Element = elmTemp.Document.Element(prmTemp.AsElementId)
                                    Dim valueElm As DB.Element = elmTemp.Document.GetElement(prmTemp.AsElementId)
                                    If valueElm IsNot Nothing Then
                                        sParamValue = valueElm.Name
                                    End If

                                Case DB.StorageType.Integer
                                    Dim sVal As Integer = prmTemp.AsInteger
                                    sParamValue = sVal.ToString

                                Case DB.StorageType.String
                                    sParamValue = prmTemp.AsString
                                    sh.Cells(iRow, iCol).NumberFormatLocal = "@"

                            End Select
                            If String.IsNullOrEmpty(sParamValue) = False Then
                                sh.Cells(iRow, iCol).Value = sParamValue
                                'sh.Cells(iRow, iCol).Select()
                            End If
                        End If

                    End If

                    '次のパラメータ
                    iCol += 1
                    sParamName = sh.Cells(1, iCol).Value
                Loop

            Catch ex As Exception

            End Try

            iRow += 1
        Next
        sh.Cells(2, 1).select()
        Return iRow - 1

    End Function

    ''' <summary>
    ''' ワークシート全体をフォーマットする
    ''' </summary>
    ''' <param name="wb"></param>
    ''' <remarks></remarks>
    Public Sub FormatWorksheet(ByVal wb As XLS.Workbook)

        'フォーマット
        With wb.Styles("標準")
            .IncludeNumber = True
            .IncludeFont = True
            .IncludeAlignment = True
            .IncludeBorder = True
            .IncludePatterns = True
            .IncludeProtection = True
            .HorizontalAlignment = XLS.Constants.xlLeft
            .VerticalAlignment = XLS.Constants.xlTop
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
        End With

        With wb.Styles("標準").Font
            .Name = "ＭＳ Ｐ明朝"
            .Size = 9
            .Bold = False
            .Italic = False
            .Underline = XLS.XlUnderlineStyle.xlUnderlineStyleNone
            .Strikethrough = False
            .ColorIndex = XLS.Constants.xlAutomatic
        End With
        wb.Styles("標準").NumberFormat = "@"

        wb.ActiveSheet.Rows.RowHeight = 35.0# '一旦すべての行の高さを35にする

    End Sub

    ''' <summary>
    ''' タイトル行をフォーマットする
    ''' </summary>
    ''' <param name="iRow"></param>
    ''' <param name="iCol"></param>
    ''' <param name="Title"></param>
    ''' <param name="Wid"></param>
    ''' <param name="sh"></param>
    ''' <remarks></remarks>
    Public Sub FormatTitleCell(ByVal iRow As Integer,
                               ByVal iCol As Integer,
                               ByVal Title As String,
                               ByVal Wid As Single,
                               ByVal sh As XLS.Worksheet)

        With sh.Cells(iRow, iCol)
            '(1)セルにタイトルを記入
            .value = Title
            '(2)上下左右の中央に配置する
            .HorizontalAlignment() = XLS.Constants.xlCenter
            .VerticalAlignment = XLS.Constants.xlCenter
            '(3)フォントの設定
            '.font.Name = "ＭＳ Ｐゴシック"
            '.font.FontStyle = "標準"
            .font.Size = 10.5
            .font.ColorIndex = 2
            '(4)セルの色
            .Interior.ColorIndex = 54
            .Interior.Pattern = XLS.Constants.xlSolid
            .Interior.PatternColorIndex = XLS.Constants.xlAutomatic
            '(5)行の幅
            .ColumnWidth = Wid
        End With
    End Sub

    ''' <summary>
    ''' エクセルのタイトル欄をフォーマットする
    ''' </summary>
    ''' <param name="iRow"></param>
    ''' <param name="iCol"></param>
    ''' <param name="Title"></param>
    ''' <param name="Wid"></param>
    ''' <param name="sh"></param>
    ''' <remarks></remarks>
    Public Sub FormatParamTitleCell(ByVal iRow As Integer,
                               ByVal iCol As Integer,
                               ByVal Title As String,
                               ByVal Wid As Single,
                               ByVal sh As XLS.Worksheet)

        With sh.Cells(iRow, iCol)
            '(1)セルにタイトルを記入
            .value = Title
            '(2)上下左右の中央に配置する
            .HorizontalAlignment() = XLS.Constants.xlCenter
            .VerticalAlignment = XLS.Constants.xlCenter
            '(3)フォントの設定
            '.font.Name = "ＭＳ Ｐゴシック"
            '.font.FontStyle = "標準"
            .font.Size = 10.5
            .font.ColorIndex = 54
            '(4)セルの色
            '.Interior.ColorIndex = 54
            '.Interior.Pattern = XLS.Constants.xlSolid
            '.Interior.PatternColorIndex = XLS.Constants.xlAutomatic
            '(5)行の幅
            .ColumnWidth = Wid
        End With
    End Sub

    ''' <summary>
    ''' エレメントのパラメータの値をエクセルに出力する
    ''' </summary>
    ''' <param name="ElementList"></param>
    ''' <param name="ParameterList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExportElementParameterValuesIT(ByVal ElementList As List(Of DB.Element),
                                                 ByVal ParameterList As List(Of String),
                                                 Optional ByVal SheetName As String = "",
                                                 Optional ByVal isNewWorkbook As Boolean = True) As Integer

        Dim wb As XLS.Workbook = GetExcelWorkBook(isNewWorkbook)
        Dim sh As XLS.Worksheet = wb.Worksheets.Add
        With wb.Application.ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
            .FreezePanes = True
        End With
        If SheetName <> "" Then
            sh.Name = SheetName
        End If

        Dim iHead As Integer = ExportParamHeader(sh, ParameterList)
        Dim iLastRow As Integer = ExportElementIT(sh, ElementList, ParameterList)

        Return iLastRow

    End Function
    ''' <summary>
    ''' エレメントのパラメータの値をエクセルに出力する
    ''' </summary>
    ''' <param name="ElementList"></param>
    ''' <param name="ParameterList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExportElementParameterValues(ByVal ElementList As List(Of DB.Element),
                                                 ByVal ParameterList As List(Of String),
                                                 Optional ByVal SheetName As String = "",
                                                 Optional ByVal isNewWorkbook As Boolean = True) As Integer

        Dim wb As XLS.Workbook = GetExcelWorkBook(isNewWorkbook)
        Dim sh As XLS.Worksheet = wb.Worksheets.Add
        With wb.Application.ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
            .FreezePanes = True
        End With
        If SheetName <> "" Then
            sh.Name = SheetName
        End If

        Dim iHead As Integer = ExportParamHeader(sh, ParameterList)
        Dim iLastRow As Integer = ExportElement(sh, ElementList)

        Return iLastRow

    End Function

    ''' <summary>
    ''' 現在アクティブなエクセルシートを読んで、対応するUIDの属性を変更する
    ''' ファミリシンボルの場合、UIDの示すファミリシンボルと記入されているシンボル名が異なる場合、複製して新たに登録する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ImportExcelAndDupricateType(ByVal dbDoc As DB.Document) As Integer

        '現在のシートを取得
        Dim wb As XLS.Workbook = GetCurrentExcelWorkBook()
        If wb Is Nothing Then
            Return -1
        End If
        Dim sh As XLS.Worksheet = wb.ActiveSheet
        If sh Is Nothing Then
            Return -2
        End If
        '1行目はタイトルであり、1:UID 2:ファミリ 3:タイプ は固定
        '2行目からUIDの値が空白になるまで読み
        Dim iRow As Integer = 2
        Dim sUID As String = sh.Cells(iRow, 1).Value
        Do While String.IsNullOrEmpty(sUID) = False
            'ファミリ名とタイプ名
            Dim sFamName As String = sh.Cells(iRow, 2).Value
            Dim sTypeName As String = sh.Cells(iRow, 3).Value
            'エレメントを取得
            Try
                'Dim elmTemp As DB.Element = dbDoc.Element(sUID)
                Dim elmTemp As DB.Element = dbDoc.GetElement(sUID)
                If elmTemp IsNot Nothing Then
                    'これがファミリシンボルであり、タイプ名が異なれば複製する
                    If TypeOf elmTemp Is DB.FamilySymbol Then
                        Dim fsmTemp As DB.FamilySymbol = TryCast(elmTemp, DB.FamilySymbol)
                        '名前は同じか?
                        If fsmTemp.Name <> sTypeName Then
                            '異なる場合は複製する
                            elmTemp = fsmTemp.Duplicate(sTypeName)
                        End If
                    End If
                End If
                'パラメータの変更
                Dim iCol As Integer = 4
                Dim ParamName As String = sh.Cells(1, iCol).Value
                Do While String.IsNullOrEmpty(ParamName) = False
                    Try
                        Dim prmTemp As DB.Parameter = elmTemp.LookupParameter(ParamName)
                        If prmTemp IsNot Nothing Then
                            If prmTemp.IsReadOnly = False Then
                                Dim sVal As String = sh.Cells(iRow, iCol).Value
                                Select Case prmTemp.StorageType
                                    Case DB.StorageType.Double
                                        If IsNumeric(sVal) = True Then
                                            prmTemp.SetValueString(sVal)
                                        End If
                                    Case DB.StorageType.Integer
                                        If IsNumeric(sVal) = True Then
                                            prmTemp.SetValueString(sVal)
                                        End If
                                    Case DB.StorageType.String
                                        prmTemp.Set(sVal)
                                End Select
                            End If
                        End If

                    Catch ex As Exception

                    End Try
                    'パラメータの取得
                    iCol += 1
                    ParamName = sh.Cells(1, iCol).Value

                Loop

            Catch ex As Exception

            End Try
            iRow += 1
            sUID = sh.Cells(iRow, 1).Value
        Loop

        Return iRow - 1

    End Function

    ''' <summary>
    ''' マテリアルを新しいエクセルブックに出力する
    ''' </summary>
    ''' <param name="lstElement"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExportElementMaterials(ByVal lstElement As List(Of DB.Element)) As Integer

        Dim wb As XLS.Workbook = GetExcelWorkBook()
        If wb Is Nothing Then
            Return -1
        End If
        Dim sh As XLS.Worksheet = wb.ActiveSheet
        If sh Is Nothing Then
            sh = wb.Worksheets.Add
        End If

        'タイトル作成
        FormatTitleCell(1, 1, "要素ID", 10, sh)
        FormatTitleCell(1, 2, "要素前", 10, sh)
        FormatTitleCell(1, 3, "カテゴリ", 10, sh)
        FormatTitleCell(1, 4, "マテリアル名", 10, sh)
        FormatTitleCell(1, 5, "面積", 10, sh)
        FormatTitleCell(1, 6, "体積", 10, sh)
        FormatTitleCell(1, 7, "説明", 10, sh)
        FormatTitleCell(1, 8, "製造元", 10, sh)
        FormatTitleCell(1, 9, "モデル", 10, sh)
        FormatTitleCell(1, 10, "価格", 10, sh)

        'タイトル行を固定
        With wb.Application.ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
            .FreezePanes = True
        End With

        '書き出し
        Dim iRow As Integer = 2
        For Each elmTemp As DB.Element In lstElement
            'sh.Cells(iRow, 1).Select()
            'これがマテリアルを所有しているかTRTする
            Try
                'Dim mSet As DB.MaterialSet = elmTemp.Materials
                Dim midSet As List(Of DB.ElementId) = elmTemp.GetMaterialIds(True)
                If midSet.Count = 0 Then
                    Continue For
                End If

                Dim ElmID As String = elmTemp.Id.IntegerValue.ToString
                Dim Name As String = elmTemp.Name
                Dim CatName As String = elmTemp.Category.Name

                'マテリアルを出力
                For Each matid As DB.ElementId In midSet
                    '体積と面積
                    Dim volTemp As Double = elmTemp.GetMaterialVolume(matid)
                    Dim areTemp As Double = elmTemp.GetMaterialArea(matid, True)

                    Dim mat As DB.Material = elmTemp.Document.GetElement(matid)
                    'マテリアルの説明(ALL_MODEL_DESCRIPTION)製造元(ALL_MODEL_MANUFACTURER)
                    'モデル(ALL_MODEL_MODEL)価格(DOOR_COST)
                    Dim prpDesc As DB.Parameter = mat.Parameter(DB.BuiltInParameter.ALL_MODEL_DESCRIPTION)
                    Dim strDesc As String = GetParamString(mat.Parameter(DB.BuiltInParameter.ALL_MODEL_DESCRIPTION))
                    Dim strManu As String = GetParamString(mat.Parameter(DB.BuiltInParameter.ALL_MODEL_MANUFACTURER))
                    Dim strModl As String = GetParamString(mat.Parameter(DB.BuiltInParameter.ALL_MODEL_MODEL))
                    Dim strCost As String = GetParamString(mat.Parameter(DB.BuiltInParameter.ALL_MODEL_COST))

                    '書き出し
                    sh.Cells(iRow, 1).Value = ElmID
                    sh.Cells(iRow, 2).Value = Name
                    sh.Cells(iRow, 3).Value = CatName
                    sh.Cells(iRow, 4).Value = mat.Name
                    sh.Cells(iRow, 5).Value = areTemp / (1000.0 / 304.8) ^ 2
                    sh.Cells(iRow, 6).Value = volTemp / (1000.0 / 304.8) ^ 3
                    sh.Cells(iRow, 7).Value = strDesc
                    sh.Cells(iRow, 8).Value = strManu
                    sh.Cells(iRow, 9).Value = strModl
                    sh.Cells(iRow, 9).Value = strCost
                    iRow += 1

                Next

                'Dim mtsitr As DB.MaterialSetIterator = mSet.ForwardIterator
                'Do While mtsitr.MoveNext
                '    Dim mat As DB.Material = mtsitr.Current
                '    '体積と面積
                '    Dim volTemp As Double = elmTemp.GetMaterialVolume(mat)
                '    Dim areTemp As Double = elmTemp.GetMaterialArea(mat)

                '    'マテリアルの説明(ALL_MODEL_DESCRIPTION)製造元(ALL_MODEL_MANUFACTURER)
                '    'モデル(ALL_MODEL_MODEL)価格(DOOR_COST)
                '    Dim prpDesc As DB.Parameter = mat.Parameter(DB.BuiltInParameter.ALL_MODEL_DESCRIPTION)
                '    Dim strDesc As String = GetParamString(mat.Parameter(DB.BuiltInParameter.ALL_MODEL_DESCRIPTION))
                '    Dim strManu As String = GetParamString(mat.Parameter(DB.BuiltInParameter.ALL_MODEL_MANUFACTURER))
                '    Dim strModl As String = GetParamString(mat.Parameter(DB.BuiltInParameter.ALL_MODEL_MODEL))
                '    Dim strCost As String = GetParamString(mat.Parameter(DB.BuiltInParameter.ALL_MODEL_COST))

                '    '書き出し
                '    sh.Cells(iRow, 1).Value = ElmID
                '    sh.Cells(iRow, 2).Value = Name
                '    sh.Cells(iRow, 3).Value = CatName
                '    sh.Cells(iRow, 4).Value = mat.Name
                '    sh.Cells(iRow, 5).Value = areTemp / (1000.0 / 304.8) ^ 2
                '    sh.Cells(iRow, 6).Value = volTemp / (1000.0 / 304.8) ^ 3
                '    sh.Cells(iRow, 7).Value = strDesc
                '    sh.Cells(iRow, 8).Value = strManu
                '    sh.Cells(iRow, 9).Value = strModl
                '    sh.Cells(iRow, 9).Value = strCost
                '    iRow += 1
                'Loop

            Catch ex As Exception

            End Try
        Next
        sh.Cells(2, 1).Select()

        Return iRow - 1

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="prmTemp"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetParamString(ByVal prmTemp As DB.Parameter) As String

        If prmTemp Is Nothing Then
            Return ""
        End If

        Dim strTemp As String = ""

        Select Case prmTemp.StorageType

            Case DB.StorageType.Double, DB.StorageType.Integer

                strTemp = prmTemp.AsValueString

            Case DB.StorageType.String

                strTemp = prmTemp.AsString

            Case Else

                strTemp = ""

        End Select

        If strTemp Is Nothing Then
            strTemp = ""
        End If

        Return strTemp

    End Function


    Public Function GetSelectedGuid() As String
        '現在のシートを取得
        Dim wb As XLS.Workbook = GetCurrentExcelWorkBook(False)
        If wb Is Nothing Then
            Return -1
        End If
        Dim sh As XLS.Worksheet = wb.ActiveSheet
        If sh Is Nothing Then
            Return -2
        End If
        '現在選択されているセル
        Dim curCell As XLS.Range = sh.Application.ActiveCell

        Return sh.Cells(curCell.Row, 1).Value



    End Function

    Public Function GetSelectedGuids() As List(Of String)

        '戻り値の準備
        Dim lstGuids As New List(Of String)

        '現在のシートを取得
        Dim wb As XLS.Workbook = GetCurrentExcelWorkBook(False)
        If wb Is Nothing Then
            Return lstGuids
        End If
        Dim sh As XLS.Worksheet = wb.ActiveSheet
        If sh Is Nothing Then
            Return lstGuids
        End If
        '現在選択されている範囲を取得する
        Dim curRange As XLS.Range = sh.Application.Selection
        For Each row1 As XLS.Range In curRange.Rows
            Dim rowNum As Integer = row1.Row
            '先頭列の値
            Dim strUid As String = sh.Cells(row1.Row, 1).Value
            If IsNothing(strUid) = True Then
                Continue For
            End If
            lstGuids.Add(strUid)
        Next
        sh = Nothing
        wb = Nothing
        Return lstGuids



    End Function

End Class

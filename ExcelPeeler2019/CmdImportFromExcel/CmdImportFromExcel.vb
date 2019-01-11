Imports Autodesk.Revit
Imports XLS = Microsoft.Office.Interop.Excel
Imports ExcelPeeler2019.My.Resources

<Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)> _
Public Class CmdImportFromExcel
    Implements UI.IExternalCommand

    ''' <summary>
    ''' エクセルから値を読み取り、Revitのパラメータの値を更新する
    ''' </summary>
    ''' <param name="commandData"></param>
    ''' <param name="message"></param>
    ''' <param name="elements"></param>
    ''' <returns></returns>
    Public Function Execute(ByVal commandData As Autodesk.Revit.UI.ExternalCommandData, ByRef message As String, ByVal elements As Autodesk.Revit.DB.ElementSet) As Autodesk.Revit.UI.Result Implements Autodesk.Revit.UI.IExternalCommand.Execute

        '******************************************
        'エクセルを取得
        '現在起動しているExcelオブジェクトを取得する
        '******************************************

        Dim exMan As New ClsExcelManager()

        Dim xlApp As XLS.Application = Nothing
        Dim xlWb As XLS.Workbook = Nothing

        If exMan.GetExcelApp(xlApp, xlWb) <> 1 Then
            message = "Excelシートを特定できません。"
            Return UI.Result.Cancelled
        End If

        'アプリケーションは解放
        exMan.ReleaseComObject(xlApp)

        '現在のシート
        Dim xlWs As XLS.Worksheet = xlWb.ActiveSheet
        exMan.ReleaseComObject(xlWb)

        If xlWs Is Nothing Then
            message = "Excelのシートを特定できません"
            Return UI.Result.Cancelled
        End If

        '******************************************
        '第一行目からパラメータ名のリストを作成
        '******************************************
        Dim lstParamNames As New Dictionary(Of String, Integer)
        Dim iCol As Integer = 2
        Dim strParamName As String = xlWs.Cells(1, iCol).Value
        strParamName = Trim(strParamName)
        Do While strParamName <> ""
            lstParamNames.Add(strParamName, iCol)
            iCol += 1
            strParamName = xlWs.Cells(1, iCol).Value
            strParamName = Trim(strParamName)
        Loop

        If lstParamNames.Count = 0 Then
            message = "タイトル行が不明です"
            exMan.ReleaseComObject(xlWs)
            Return UI.Result.Cancelled
        End If

        '******************************************
        '第2行目から値を部屋に設定、成功した行は削除
        '******************************************
        Dim uiApp As UI.UIApplication = commandData.Application
        Dim uiDoc As UI.UIDocument = uiApp.ActiveUIDocument
        Dim dbDoc As DB.Document = uiDoc.Document
        Dim iRow As Integer = 2
        Dim strUID As String = xlWs.Cells(iRow, 1).Value
        Dim lstElmInGroup As New List(Of DB.Element)
        Dim uiRes As UI.Result = UI.Result.Cancelled
        Using tr1 As New DB.Transaction(dbDoc, CMD_IMPORTFROMEXCEL)
            If tr1.Start = DB.TransactionStatus.Started Then
                Try
                    Do While strUID <> ""
                        xlWs.Cells(iRow, 1).Select()
                        'UIDをエレメントに変換
                        'Dim elm As DB.Element = dbDoc.Element(strUID)
                        Dim elm As DB.Element = dbDoc.GetElement(strUID)
                        If elm Is Nothing Then
                            iRow += 1
                            strUID = xlWs.Cells(iRow, 1).Value
                            Continue Do
                        End If

                        'グループに属していたら編集できないので、これをためておく
                        If elm.GroupId <> DB.ElementId.InvalidElementId Then
                            lstElmInGroup.Add(elm)
                            iRow += 1
                            strUID = xlWs.Cells(iRow, 1).Value
                            Continue Do
                        End If

                        'パラメータを設定
                        For Each strParam As String In lstParamNames.Keys
                            Try
                                Dim prm As DB.Parameter = elm.LookupParameter(strParam)
                                If IsNothing(prm) = True Then
                                    Continue For
                                End If
                                '読み取り専用のパラメータはのぞく
                                If prm.IsReadOnly = True Then
                                    'なにもしない
                                Else
                                    Dim curCell As XLS.Range = xlWs.Cells(iRow, lstParamNames.Item(strParam))
                                    curCell.Select()
                                    Dim strValue As String = curCell.Text
                                    strValue = ExcelText(strValue)
                                    If prm.StorageType = DB.StorageType.String Then
                                        prm.Set(strValue)
                                        curCell.Interior.ColorIndex = 15
                                        curCell.Interior.Pattern = XLS.Constants.xlSolid
                                    ElseIf prm.StorageType = DB.StorageType.Double Then
                                        prm.SetValueString(strValue)
                                        curCell.Interior.ColorIndex = 15
                                        curCell.Interior.Pattern = XLS.Constants.xlSolid
                                    ElseIf prm.StorageType = DB.StorageType.Integer Then
                                        If IsNumeric(strValue) = True Then
                                            prm.Set(Integer.Parse(strValue))
                                            curCell.Interior.ColorIndex = 15
                                            curCell.Interior.Pattern = XLS.Constants.xlSolid
                                        End If
                                    End If
                                End If

                            Catch ex As Exception

                            End Try

                        Next
                        iRow += 1
                        strUID = xlWs.Cells(iRow, 1).Value

                    Loop
                    xlWs.Cells(1, 1).Select()
                    dbDoc.Regenerate()

                    tr1.Commit()
                    uiRes = UI.Result.Succeeded

                Catch ex As Exception
                    message = ex.Message
                    tr1.RollBack()
                    uiRes = UI.Result.Cancelled
                End Try
            End If
        End Using


        If uiRes = UI.Result.Succeeded Then
            If lstElmInGroup.Count = 0 Then
                MsgBox("読み込みが終了しました。成功したセルは灰色になっています。", MsgBoxStyle.OkOnly, CMD_IMPORTFROMEXCEL)
            Else
                MsgBox("読み込みが終了しました。成功したセルは灰色になっていますが、" + lstElmInGroup.Count.ToString + "の要素はグループに属しているため変更されませんでした。", MsgBoxStyle.OkOnly, CMD_IMPORTFROMEXCEL)
            End If
        End If

        Return uiRes



        Return UI.Result.Succeeded




    End Function

    ''' <summary>
    ''' エクセルから取得した文字列はNothingの場合がある。また
    ''' </summary>
    ''' <param name="strTemp"></param>
    ''' <returns></returns>
    Function ExcelText(ByVal strTemp As String) As String

        strTemp = Trim(strTemp)
        If String.IsNullOrEmpty(strTemp) Then
            Return ""
        Else
            Dim newStr As String = ""
            Dim c As Integer = 1
            For c = 1 To strTemp.Length
                Dim ch As String = Mid(strTemp, c, 1)
                If Asc(ch) = 10 Then
                    '一つ前に13がないことが条件
                    If c <= 1 Then
                        newStr = newStr + vbCrLf
                    ElseIf Asc(Mid(strTemp, c - 1, 1)) <> 13 Then
                        newStr = newStr + vbCrLf
                    Else
                        newStr = newStr + ch
                    End If
                Else
                    newStr = newStr + ch
                End If
            Next
            Return newStr

        End If


    End Function
End Class

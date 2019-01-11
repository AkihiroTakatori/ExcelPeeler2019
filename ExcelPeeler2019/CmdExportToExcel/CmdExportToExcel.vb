Imports Autodesk.Revit
Imports ExcelPeeler2019.My.Resources

<Attributes.Transaction(Attributes.TransactionMode.ReadOnly)>
Public Class CmdExportToExcel
    Implements UI.IExternalCommand

    Public Function Execute(commandData As Autodesk.Revit.UI.ExternalCommandData, ByRef message As String, elements As Autodesk.Revit.DB.ElementSet) As Autodesk.Revit.UI.Result Implements Autodesk.Revit.UI.IExternalCommand.Execute

        '************************************************************************
        '
        'タスクダイアログで出力範囲を選択
        '
        '************************************************************************
        Dim uiDoc As UI.UIDocument = commandData.Application.ActiveUIDocument
        Dim tsk1 As New UI.TaskDialog(CMD_EXPORTTOEXCEL)
        With tsk1
            .TitleAutoPrefix = False
            .MainInstruction = "書出し範囲の選択"
            .MainContent = "書出す要素の範囲を選択してください。"
            .AddCommandLink(UI.TaskDialogCommandLinkId.CommandLink1, "プロジェクト全体の要素")
            .AddCommandLink(UI.TaskDialogCommandLinkId.CommandLink2, "現在のビューにある要素")
            .AddCommandLink(UI.TaskDialogCommandLinkId.CommandLink3, "現在選択されている要素")
            .CommonButtons = UI.TaskDialogCommonButtons.Cancel
            .DefaultButton = UI.TaskDialogResult.CommandLink1
        End With

        Dim res As UI.TaskDialogResult = tsk1.Show
        Dim mode As Integer = 0
        If res = UI.TaskDialogResult.CommandLink1 Then
            mode = 0
        ElseIf res = UI.TaskDialogResult.CommandLink2 Then
            mode = 1
        ElseIf res = UI.TaskDialogResult.CommandLink3 Then
            mode = 2
            '選択されていない場合は終了
            If uiDoc.Selection.GetElementIds.Count = 0 Then
                MsgBox("要素を選択してから起動してください。", MsgBoxStyle.OkOnly, CMD_EXPORTTOEXCEL)
                Return UI.Result.Cancelled
            End If
        Else
            Return UI.Result.Cancelled
        End If

        '************************************************************************
        '
        'ダイアログでパラメータを選択
        '
        '************************************************************************
        Dim dlg1 As New dlgCatParam(commandData, mode)
        If dlg1.ShowDialog <> Windows.Forms.DialogResult.OK Then
            Return UI.Result.Cancelled
        End If

        MsgBox("出力が完了しました。", MsgBoxStyle.OkOnly, CMD_EXPORTTOEXCEL)

        Return UI.Result.Succeeded

    End Function
End Class

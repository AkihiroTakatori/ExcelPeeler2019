Imports System
Imports Autodesk
Imports Autodesk.Revit
Imports Autodesk.Revit.UI
Imports ExcelPeeler2019.My.Resources

''' <summary>
''' リボンの追加
''' </summary>
<Autodesk.Revit.Attributes.Transaction(Attributes.TransactionMode.Manual)>
Public Class AppAddRibbon
    Implements Autodesk.Revit.UI.IExternalApplication

    Public Function OnStartup(application As UIControlledApplication) As Result Implements IExternalApplication.OnStartup

        'このdllへののパスを作成する
        Dim strMyPath As String = IO.Path.Combine(My.Application.Info.DirectoryPath)


        'タブがすでに追加されている場合はパネルを取得
        Dim targetPanel As UI.RibbonPanel = Nothing
        Dim lstRbnPnls As List(Of UI.RibbonPanel) = application.GetRibbonPanels(RBN_TAB_NAME)
        If lstRbnPnls.Count > 0 Then
            For Each RbnPnl0 As UI.RibbonPanel In lstRbnPnls
                If RbnPnl0.Name = RBN_PNL_NAME Then
                    targetPanel = RbnPnl0
                    Exit For
                End If
            Next
        End If

        'パネルが見つからなかった場合は


        'ボタンデータを作成する





    End Function

    Public Function OnShutdown(application As UIControlledApplication) As Result Implements IExternalApplication.OnShutdown
        Return Result.Succeeded
    End Function
End Class

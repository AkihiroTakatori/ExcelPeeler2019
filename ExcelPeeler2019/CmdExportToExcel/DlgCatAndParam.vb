Imports System.Windows.Forms
Imports System.IO
Imports Autodesk.Revit

Public Class dlgCatParam

    Dim m_dbDoc As DB.Document
    Dim m_uiDoc As UI.UIDocument


    ''' <summary>
    ''' 設定ファイルのパス
    ''' </summary>
    Dim m_PthDef As String

    ''' <summary>
    ''' パラメータを出力する対象の要素
    ''' </summary>
    Dim m_lstElement As New List(Of DB.Element)

    ''' <summary>
    ''' 出力するパラメータのリスト
    ''' </summary>
    Dim m_lstParamName As New List(Of String)

    ''' <summary>
    ''' ファイルから読み込んだか？設定ファイルには出力順も指定されているのでパラメータ名のリストは作成しない
    ''' </summary>
    Dim m_bFromFile As Boolean = False

    ''' <summary>
    ''' 選択する範囲
    ''' </summary>
    Dim m_SelectMode As Integer
    Dim m_SelIds As List(Of DB.ElementId)


    ''' <summary>
    ''' OKボタンのアクション
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        '書き出すパラメータの確認
        If ChkbxParameters.CheckedItems.Count = 0 Then
            MsgBox("No parameter is checked.", MsgBoxStyle.OkOnly, Me.Text)
            Exit Sub
        End If

        'チェックボックスの状態を保存する
        Dim iCatItem As ItmCategory = TryCast(LbxCategories.SelectedItem, ItmCategory)
        Dim defFileName As String = DefFilePath(iCatItem.Category.Name)

        '値の書き出し
        Dim lstChkIdx As New List(Of String)
        For Each idx As Integer In ChkbxParameters.CheckedIndices
            lstChkIdx.Add(idx.ToString)
        Next
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")
        Try
            System.IO.File.WriteAllLines(defFileName, lstChkIdx, enc)
        Catch ex As Exception
        End Try

        'パラメータ名称のリスト
        If m_bFromFile = False Then
            m_lstParamName.Clear()
            For Each itm As Object In ChkbxParameters.CheckedItems
                Dim strPname As String = itm.ToString
                m_lstParamName.Add(strPname)
            Next
        End If

        Me.Hide()

        Dim bType As Boolean
        If RbnType.Checked = True Then
            bType = True
        Else
            bType = False
        End If
        Dim dlg1 As New dlgCatParamExp(IO.Path.GetFileNameWithoutExtension(defFileName), m_lstElement, m_lstParamName)
        If dlg1.ShowDialog() <> Windows.Forms.DialogResult.OK Then
            Me.Show()
            Exit Sub
        End If

        '現在のダイアログボックスの値を保存する
        SaveLastValue(LbxCategories, Me.Name)
        SaveLastValue(RbnType, Me.Name)
        SaveLastValue(RbnInstance, Me.Name)


        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    ''' <summary>
    ''' ダイアログボックスの新規作成
    ''' </summary>
    ''' <param name="CmdData"></param>
    ''' <param name="SelectMode"></param>
    Public Sub New(ByVal CmdData As UI.ExternalCommandData, Optional ByVal SelectMode As Integer = 0)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Me.Text = My.Resources.CMD_EXPORTTOEXCEL

        '選択範囲
        m_SelectMode = SelectMode

        m_uiDoc = CmdData.Application.ActiveUIDocument
        m_dbDoc = m_uiDoc.Document

        '定義ファイルの位置(MyDocument/Revit Peeler)
        m_PthDef = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        m_PthDef = IO.Path.Combine(m_PthDef, My.Application.Info.CompanyName, My.Application.Info.ProductName)

        If IO.Directory.Exists(m_PthDef) = False Then
            Try
                IO.Directory.CreateDirectory(m_PthDef)
            Catch ex As Exception
            End Try
        End If

        If SelectMode = 2 Then
            m_SelIds = m_uiDoc.Selection.GetElementIds
            If m_SelIds.Count = 0 Then
                Err.Raise(vbObjectError + 513, Me.Text, "なにも選択されていません。")
                Exit Sub
            End If
        End If


        'カテゴリのリストを作成し、リストボックスを充てんする(モードに応じてあるものだけを選択)
        For Each cat As DB.Category In m_dbDoc.Settings.Categories
            Dim catFilt As New DB.ElementCategoryFilter(cat.Id)
            Dim lstElmIds As List(Of DB.ElementId) = Nothing
            If m_SelectMode = 0 Then

                'プロジェクト全体
                Dim collctor0 As New DB.FilteredElementCollector(m_dbDoc)
                lstElmIds = collctor0.WherePasses(catFilt).ToElementIds
            ElseIf m_SelectMode = 1 Then

                '現在のビュー
                Dim collector1 As New DB.FilteredElementCollector(m_dbDoc, m_dbDoc.ActiveView.Id)
                lstElmIds = collector1.WherePasses(catFilt).ToElementIds
            Else

                '選択している要素から
                Dim collector2 As New DB.FilteredElementCollector(m_dbDoc, m_SelIds)
                lstElmIds = collector2.WherePasses(catFilt).ToElementIds
            End If

            If lstElmIds.Count > 0 Then
                Dim iCatItem As New ItmCategory(cat)
                LbxCategories.Items.Add(iCatItem)
            End If
        Next

        '前回の値のリストア
        SetLastValue(LbxCategories, Me.Name)
        SetLastValue(RbnType, Me.Name)
        SetLastValue(RbnInstance, Me.Name)

    End Sub



    ''' <summary>
    ''' カテゴリを選択したときのアクション
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub LbxCategories_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles LbxCategories.SelectedIndexChanged

        'まずリストはクリア
        ChkbxParameters.Items.Clear()
        m_lstElement.Clear()

        '結局文字列しか使用しないので、パラメータの文字列を集める仕様に変更する20111204

        'このカテゴリのタイプまたはインスタンスのパラメータを集める

        Dim collector As DB.FilteredElementCollector
        If m_SelectMode = 0 Then
            collector = New DB.FilteredElementCollector(m_dbDoc)
        ElseIf m_SelectMode = 1 Then
            '現在のビュー～選択されるのはインスタンス
            collector = New DB.FilteredElementCollector(m_dbDoc, m_dbDoc.ActiveView.Id)
        Else
            '選択しているものから～選択しているのはインスタンス
            collector = New DB.FilteredElementCollector(m_dbDoc, m_SelIds)
        End If
        'カテゴリフィルタ
        Dim iCatItem As ItmCategory = TryCast(LbxCategories.SelectedItem, ItmCategory)
        Dim catFilter As New DB.ElementCategoryFilter(iCatItem.Category.Id)

        collector.WherePasses(catFilter)
        If RbnType.Checked = True Then
            'タイプの場合
            If m_SelectMode = 0 Then
                collector.WhereElementIsElementType()
                m_lstElement = collector.ToElements
            Else
                'インスタンスからタイプのリストを作成する
                Dim lstInstanceElms As List(Of DB.Element) = collector.ToElements
                If lstInstanceElms.Count > 0 Then
                    Dim lstTypeIds As New List(Of DB.ElementId)
                    For Each insElm As DB.Element In lstInstanceElms
                        'タイプを求めてみる
                        Try
                            If lstTypeIds.Contains(insElm.GetTypeId) = False Then
                                If insElm.GetTypeId.Equals(DB.ElementId.InvalidElementId) = False Then
                                    lstTypeIds.Add(insElm.GetTypeId)
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                    Next
                    For Each typElmId As DB.ElementId In lstTypeIds
                        Dim typElm As DB.Element = m_dbDoc.GetElement(typElmId)
                        m_lstElement.Add(typElm)
                    Next
                Else
                    m_lstElement = New List(Of DB.Element)
                End If
            End If
        Else
            collector.WhereElementIsNotElementType()
            m_lstElement = collector.ToElements
        End If

        If m_lstElement.Count = 0 Then
            Exit Sub
        End If

        ProgressBar1.Maximum = m_lstElement.Count
        ProgressBar1.Value = 0
        For Each elmTemp As DB.Element In m_lstElement
            ProgressBar1.Value += 1
            ProgressBar1.Update()
            For Each prmTemp As DB.Parameter In elmTemp.Parameters

                Dim strParamName As String = prmTemp.Definition.Name
                If ChkbxParameters.Items.Contains(strParamName) = False Then
                    ChkbxParameters.Items.Add(strParamName)
                End If

            Next
        Next

        'ファミリインスタンスの場合はToRoom FromRoom Room Spaceも追加
        Dim fi As DB.FamilyInstance = TryCast(m_lstElement.Item(0), DB.FamilyInstance)
        If fi IsNot Nothing Then
            ChkbxParameters.Items.Add("ToRoom")
            ChkbxParameters.Items.Add("FromRoom")
            ChkbxParameters.Items.Add("Room")
            ChkbxParameters.Items.Add("Space")
            ChkbxParameters.Items.Add("Host")
        End If
        ChkbxParameters.Items.Add("ID")

        '記録用定義ファイルを検索してチェックボックスの状態を設定する
        '定義ファイルはカテゴリ名+タイプまたはインスタンス.txt
        Dim defFileName As String = DefFilePath(iCatItem.Category.Name)
        If IO.File.Exists(defFileName) = False Then
            Exit Sub
        End If

        '定義ファイルにはチェックされたインデックスの番号が並んでいる
        Dim strIndex As String() = File.ReadAllLines(defFileName)
        For Each str1 As String In strIndex
            If IsNumeric(str1) = True Then
                Dim idx As Integer = CInt(str1)
                If idx < ChkbxParameters.Items.Count Then
                    Try
                        'インデックスの範囲に入っていないかもしれないのでtrycatch
                        ChkbxParameters.SetItemChecked(idx, True)
                    Catch ex As Exception
                    End Try
                End If
            End If
        Next



    End Sub

    Private Sub RbnType_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RbnType.CheckedChanged
        LbxCategories_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub MnuAllCheck_Click(sender As System.Object, e As System.EventArgs) Handles mnuAllCheck.Click

        For i As Integer = 0 To ChkbxParameters.Items.Count - 1
            ChkbxParameters.SetItemChecked(i, True)
        Next


    End Sub

    Private Sub MnuAllUnCheck_Click(sender As System.Object, e As System.EventArgs) Handles mnuAllUnCheck.Click
        For i As Integer = 0 To ChkbxParameters.Items.Count - 1
            ChkbxParameters.SetItemChecked(i, False)
        Next

    End Sub

    Private Function DefFilePath(ByVal CatName As String) As String

        '定義ファイルはカテゴリ名+タイプまたはインスタンス.txt
        Dim defFileName As String = ""
        If RbnType.Checked = True Then
            defFileName = CatName + "TYPE.txt"
        Else
            defFileName = CatName + "INSTANCE.txt"
        End If
        Return IO.Path.Combine(m_PthDef, defFileName)

    End Function

    ''' <summary>
    ''' 設定ファイルをロードして、次のダイアログボックスに進む
    ''' ------------------------------------------------------
    ''' ファイルの形式は*.txtで
    ''' (1)カテゴリ名TYPE/INSTANCE
    ''' (2)インスタンスorタイプ
    ''' (3)パラメーター名称（順番に並んでカンマ区切り)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnLoad_Click(sender As System.Object, e As System.EventArgs) Handles BtnLoad.Click

        'ファイルを開くダイアログ
        Dim dlgFileOpen As New OpenFileDialog()
        With dlgFileOpen
            .Title = "パラメーター設定ファイルを読み込み"
            .Filter = "パラメーター設定ファイル(*.txt)|*.txt"
            .CheckFileExists = True
        End With

        If dlgFileOpen.ShowDialog <> Windows.Forms.DialogResult.OK Then
            Exit Sub
        End If

        'ファイルを開いて読み込む。
        '適正な形式であることを確認する。
        Dim strFileName As String = dlgFileOpen.FileName
        Try
            '読み込み
            Dim strLines As String() = System.IO.File.ReadAllLines(strFileName)
            '1行目はカテゴリ名TYPE/INSTANCE
            Dim strCatNameTI As String = strLines(0)
            Dim strCatName As String = ""
            If strCatNameTI.EndsWith("TYPE") = True Then
                strCatName = strCatNameTI.Remove(strCatNameTI.Length - 4)
                RbnType.Checked = True
                RbnInstance.Checked = False
            ElseIf strCatNameTI.EndsWith("INSTANCE") = True Then
                strCatName = strCatNameTI.Remove(strCatNameTI.Length - 8)
                RbnType.Checked = False
                RbnInstance.Checked = True
            Else
                MsgBox("選択されたファイルは適切ではありません。", MsgBoxStyle.OkOnly, Me.Text)
                Exit Sub
            End If
            'これがリストに存在するか？
            Dim iCatecory As Integer = -1
            For i As Integer = 0 To LbxCategories.Items.Count - 1
                Dim iC As ItmCategory = LbxCategories.Items(i)
                If UCase(iC.ToString) = UCase(strCatName) Then
                    iCatecory = i
                    Exit For
                End If
            Next
            '何も選択されなかった場合
            If iCatecory = -1 Then
                MsgBox("選択されたファイルは適切ではありません。", MsgBoxStyle.OkOnly, Me.Text)
                Exit Sub
            End If

            LbxCategories.SelectedIndex = iCatecory
            LbxCategories_SelectedIndexChanged(Me, e)

            '2行目からはパラメーターの羅列
            Dim lstParamNameOnFile As New List(Of String)

            For j As Integer = 0 To ChkbxParameters.Items.Count - 1
                ChkbxParameters.SetItemChecked(j, False)
            Next
            m_lstParamName.Clear()
            For i As Integer = 1 To strLines.Length - 1
                Dim strPName As String = strLines(i)
                'チェックリストに存在するか？
                Dim iFind As Integer = -1
                For j As Integer = 0 To ChkbxParameters.Items.Count - 1
                    If UCase(strPName) = UCase(ChkbxParameters.Items(j)) Then
                        m_lstParamName.Add(strPName)
                        ChkbxParameters.SetItemChecked(j, True)
                        Exit For
                    End If
                Next
            Next
            If m_lstParamName.Count = 0 Then
                MsgBox("選択されたファイルは適切ではありません。", MsgBoxStyle.OkOnly, Me.Text)
                Exit Sub
            End If


            '次のダイアログへ進む
            m_bFromFile = True
            OK_Button_Click(Me, e)
            m_bFromFile = False


        Catch ex As Exception

        End Try

    End Sub



End Class

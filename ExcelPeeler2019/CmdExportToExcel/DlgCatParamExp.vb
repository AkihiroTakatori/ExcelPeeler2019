Imports System.Windows.Forms
Imports Autodesk
Imports Autodesk.Revit

Public Class dlgCatParamExp

    Dim m_PthDef As String
    Dim m_CatName As String
    Dim m_ParamList As List(Of ItmParameter)
    Dim m_ElementList As List(Of DB.Element)
    Dim m_IsType As Boolean

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click


        '値の保存
        Dim lstItm As New List(Of String)
        For Each itm As String In lbxParamSort.Items
            lstItm.Add(itm)
        Next
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")
        Try
            System.IO.File.WriteAllLines(m_PthDef, lstItm, enc)
        Catch ex As Exception
        End Try


        Dim lstParamName As New List(Of String)
        For Each itm As String In lbxParamSort.Items
            lstParamName.Add(itm)

        Next

        'エクセルを起動して値を書き出す
        Dim iExUtil As New ClsExcelUtils
        iExUtil.ExportElementParameterValues(m_ElementList, lstParamName)

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public Sub New(ByVal CategoryName As String, _
                   ByVal ElementList As List(Of DB.Element), _
                   ByVal PalameteItemrList As List(Of String))

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        m_CatName = CategoryName
        'm_ParamList = PalameteItemrList
        m_ElementList = ElementList
        If CategoryName.EndsWith("TYPE") Then
            m_IsType = True
        Else
            m_IsType = False
        End If

        '定義ファイルの位置
        m_PthDef = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        m_PthDef = m_PthDef + "\" + My.Application.Info.CompanyName + "\" + My.Application.Info.ProductName + "\"
        'CategoryName = Replace(CategoryName, ".txt", "")
        m_PthDef = m_PthDef + CategoryName + "SORT.txt"

        '定義ファイルが存在するか?
        If System.IO.File.Exists(m_PthDef) = True Then
            '存在する場合は、順番通りに読み込む
            Dim sDef As String() = IO.File.ReadAllLines(m_PthDef, System.Text.Encoding.GetEncoding("shift-jis"))
            For Each sDefItem As String In sDef

                For Each prmitm As String In PalameteItemrList
                    If prmitm = sDefItem Then
                        lbxParamSort.Items.Add(prmitm) 'リストボックスにに追加
                        PalameteItemrList.Remove(prmitm) '検索時間短縮のために候補のリストから削除
                        Exit For
                    End If
                Next
            Next
            '定義ファイルに存在しない残りを追加する
            For Each prmItem As String In PalameteItemrList
                lbxParamSort.Items.Add(prmItem)
            Next
        Else
            '定義が存在しない場合は普通に追加する
            For Each prmitm As String In PalameteItemrList
                lbxParamSort.Items.Add(prmitm)
            Next

        End If

        lbxParamSort.SelectedIndex = 0

    End Sub

    Private Sub btnUP_Click(sender As System.Object, e As System.EventArgs) Handles btnUP.Click

        If lbxParamSort.SelectedIndex <= 0 Then
            Exit Sub
        End If

        Dim iPi As String = TryCast(lbxParamSort.SelectedItem, String)
        Dim intPi As Integer = lbxParamSort.SelectedIndex

        lbxParamSort.Items.RemoveAt(intPi)
        lbxParamSort.Items.Insert(intPi - 1, iPi)
        lbxParamSort.SelectedIndex = intPi - 1

    End Sub

    Private Sub btnDn_Click(sender As System.Object, e As System.EventArgs) Handles btnDn.Click
        Dim intPi As Integer = lbxParamSort.SelectedIndex
        If intPi = lbxParamSort.Items.Count - 1 Then
            Exit Sub
        End If
        If intPi < 0 Then
            Exit Sub
        End If

        Dim iPi As String = TryCast(lbxParamSort.SelectedItem, String)
        lbxParamSort.Items.RemoveAt(intPi)
        lbxParamSort.Items.Insert(intPi + 1, iPi)
        lbxParamSort.SelectedIndex = intPi + 1
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        'ファイルを開く
        Dim dlgOpenFile As New OpenFileDialog()
        With dlgOpenFile
            .Title = "設定ファイル名を指定"
            .Filter = "パラメーター設定ファイル(*.txt)|*.txt"
            .CheckFileExists = False
        End With

        If dlgOpenFile.ShowDialog <> Windows.Forms.DialogResult.OK Then
            Exit Sub
        End If

        '書き込みコンテンツの作成
        '1行目はかとごり名TYPE カテゴリ名INSTANCE
        Dim lstContent As New List(Of String)
        lstContent.Add(m_CatName)
        For j As Integer = 0 To lbxParamSort.Items.Count - 1
            lstContent.Add(lbxParamSort.Items(j))
        Next

        '書き込み
        System.IO.File.WriteAllLines(dlgOpenFile.FileName, lstContent)


    End Sub
End Class

Imports System
Imports System.Windows.Forms
Imports Autodesk
Imports Autodesk.Revit

Module MdlDialogUtil

    Public Function SetLastValue(ByVal TextBoxControl As TextBox,
                                 ByVal DialogBoxName As String,
                                 Optional ByVal DefaultValue As String = "",
                                 Optional ByVal OnlyGetValue As Boolean = False) As String

        Dim strLast As String = GetSetting(My.Application.Info.AssemblyName, DialogBoxName, TextBoxControl.Name, DefaultValue)
        If OnlyGetValue = False Then
            TextBoxControl.Text = strLast
        End If
        Return strLast

    End Function

    Public Sub SaveLastValue(ByVal TextBoxControl As TextBox, ByVal DialogBoxName As String)
        SaveSetting(My.Application.Info.AssemblyName, DialogBoxName, TextBoxControl.Name, TextBoxControl.Text)
    End Sub

    ''' <summary>
    ''' コンボボックスに前回の値を復元する
    ''' </summary>
    ''' <param name="ComboBoxControl">コンボボックスコントロール</param>
    ''' <param name="DialogBoxName">アプリケーションの名前</param>
    ''' <param name="DefaultValue">ドロップダウンリストのときは数字の文字列を初期値として設定する。指定しないと0。</param>
    ''' <param name="OnlyGetValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SetLastValue(ByVal ComboBoxControl As ComboBox,
                                 ByVal DialogBoxName As String,
                                 Optional ByVal DefaultValue As String = "",
                                 Optional ByVal OnlyGetValue As Boolean = False) As Integer

        Dim strLast As String = GetSetting(My.Application.Info.AssemblyName, DialogBoxName, ComboBoxControl.Name, DefaultValue.ToString)
        If ComboBoxControl.DropDownStyle = ComboBoxStyle.DropDownList Then
            Dim intLast As Integer = 0
            If ComboBoxControl.Items.Count = 0 Then
                'アイテムがない場合
                intLast = -1
            Else
                'ドロップダウンリストの場合は数字ベース
                If strLast = "" Then
                    strLast = "0"
                End If
                If IsNumeric(strLast) = True Then
                    intLast = Integer.Parse(strLast)
                End If
                If intLast < 0 Then
                    intLast = 0
                ElseIf ComboBoxControl.Items.Count <= intLast Then
                    intLast = ComboBoxControl.Items.Count - 1
                End If
            End If

            If OnlyGetValue = False Then
                ComboBoxControl.SelectedIndex = intLast
            End If

            Return intLast

        Else
            'そうでなければ文字ベース
            Dim intLast As Integer = -1
            If ComboBoxControl.Items.Contains(DefaultValue) = True Then
                intLast = ComboBoxControl.Items.IndexOf(DefaultValue)
            End If
            If OnlyGetValue = False Then
                ComboBoxControl.Text = DefaultValue
            End If
            Return intLast
        End If



    End Function


    Public Sub SaveLastValue(ByVal ComboBoxControl As ComboBox, ByVal DialogBoxName As String)
        If ComboBoxControl.DropDownStyle = ComboBoxStyle.DropDownList Then
            SaveSetting(My.Application.Info.AssemblyName, DialogBoxName, ComboBoxControl.Name, ComboBoxControl.SelectedIndex.ToString)
        Else
            SaveSetting(My.Application.Info.AssemblyName, DialogBoxName, ComboBoxControl.Name, ComboBoxControl.Text)
        End If
    End Sub

    Public Function SetLastValue(ByVal ListBoxControl As ListBox, ByVal DialogBoxName As String, Optional ByVal DefaultValue As Integer = 0, Optional ByVal OnlyGetValue As Boolean = False) As Integer

        Dim strLast As String = GetSetting(My.Application.Info.AssemblyName, DialogBoxName, ListBoxControl.Name, DefaultValue.ToString)
        Dim intLast As Integer = 0
        If IsNumeric(strLast) = True Then
            intLast = Integer.Parse(strLast)
        End If
        If OnlyGetValue = False And ListBoxControl.Items.Count > 0 Then

            If -1 <= intLast And intLast < ListBoxControl.Items.Count Then
                ListBoxControl.SelectedIndex = intLast
            Else
                ListBoxControl.SelectedIndex = 0
            End If
        End If

        Return intLast


    End Function

    Public Sub SaveLastValue(ByVal ListBoxControl As ListBox, ByVal DialogBoxName As String)
        SaveSetting(My.Application.Info.AssemblyName, DialogBoxName, ListBoxControl.Name, ListBoxControl.SelectedIndex.ToString)
    End Sub

    Public Function SetLastValue(ByVal CheckBoxControl As CheckBox, ByVal DialogBoxName As String, Optional ByVal DefaultValue As Boolean = True, Optional ByVal OnlyGetValue As Boolean = False) As Boolean

        Dim strLast As String = GetSetting(My.Application.Info.AssemblyName, DialogBoxName, CheckBoxControl.Name, DefaultValue.ToString)
        Dim bolLast As Boolean = DefaultValue
        Try
            bolLast = Boolean.Parse(strLast)
        Catch ex As Exception

        End Try

        If OnlyGetValue = False Then
            CheckBoxControl.Checked = bolLast
        End If
        Return bolLast

    End Function

    Public Sub SaveLastValue(ByVal CheckBoxControl As CheckBox, ByVal DialogBoxName As String)
        SaveSetting(My.Application.Info.AssemblyName, DialogBoxName, CheckBoxControl.Name, CheckBoxControl.Checked.ToString)
    End Sub

    Public Function SetLastValue(ByVal RadioButtonControl As RadioButton, ByVal DialogBoxName As String, Optional ByVal DefaultValue As Boolean = True, Optional ByVal OnlyGetValue As Boolean = False) As Boolean

        Dim strLast As String = GetSetting(My.Application.Info.AssemblyName, DialogBoxName, RadioButtonControl.Name, DefaultValue.ToString)
        Dim bolLast As Boolean = DefaultValue
        Try
            bolLast = Boolean.Parse(strLast)
        Catch ex As Exception

        End Try

        If OnlyGetValue = False Then
            RadioButtonControl.Checked = bolLast
        End If

        Return bolLast

    End Function

    Public Sub SaveLastValue(ByVal RadioButtonControl As RadioButton, ByVal DialogBoxName As String)
        SaveSetting(My.Application.Info.AssemblyName, DialogBoxName, RadioButtonControl.Name, RadioButtonControl.Checked.ToString)
    End Sub

    Public Function SetLastValue(ByVal NumericUpDnControl As NumericUpDown,
                                 ByVal DialogBoxName As String,
                                 Optional ByVal DefaultValue As Decimal = 0,
                                 Optional ByVal OnlyGetValue As Boolean = False) As Integer
        Dim strLast As String = GetSetting(My.Application.Info.AssemblyName, DialogBoxName, NumericUpDnControl.Name, DefaultValue)
        Dim intLast As Decimal = 0
        If IsNumeric(strLast) = True Then
            intLast = Decimal.Parse(strLast)
        End If
        If intLast < NumericUpDnControl.Minimum Then
            intLast = NumericUpDnControl.Minimum
        ElseIf NumericUpDnControl.Maximum < intLast Then
            intLast = NumericUpDnControl.Maximum
        End If

        If OnlyGetValue = False Then
            NumericUpDnControl.Value = intLast
        End If

        Return intLast

    End Function

    Public Sub SaveLastValue(ByVal NumericUpDnControl As NumericUpDown,
                                 ByVal DialogBoxName As String)
        SaveSetting(My.Application.Info.AssemblyName, DialogBoxName, NumericUpDnControl.Name, NumericUpDnControl.Value.ToString)

    End Sub

    Public Sub SaveLastValue(ByVal chekedListBoxControl As CheckedListBox,
                             ByVal DialogBoxName As String)

        'インデックスをカンマ区切りで保存
        Dim strChecked As String = ""
        For Each idx As Integer In chekedListBoxControl.CheckedIndices
            If strChecked = "" Then
                strChecked = idx.ToString
            Else
                strChecked = strChecked + "," + idx.ToString
            End If
        Next
        SaveSetting(My.Application.Info.AssemblyName, DialogBoxName, chekedListBoxControl.Name, strChecked)

    End Sub

    Public Function SetLastValue(ByVal checkedListBoxCOntrol As CheckedListBox,
                            ByVal DialogBoxName As String,
                            Optional ByVal DefaultValue As String = "",
                            Optional ByVal OnlyGetValue As Boolean = False) As String


        Dim strChecked As String = GetSetting(My.Application.Info.AssemblyName, DialogBoxName, checkedListBoxCOntrol.Name, DefaultValue)
        If strChecked = "" Then
        Else
            Dim strVals() As String = Split(strChecked, ",")
            If OnlyGetValue = False Then
                For i As Integer = 0 To strVals.Length - 1
                    Dim strV As String = strVals(i)
                    If IsNumeric(strV) = False Then
                        Continue For
                    End If
                    Dim intV As Integer = Integer.Parse(strV)
                    If 0 <= intV And intV < checkedListBoxCOntrol.Items.Count Then
                        checkedListBoxCOntrol.SetItemChecked(intV, True)
                    End If
                Next
            End If
        End If

        Return strChecked

    End Function


End Module

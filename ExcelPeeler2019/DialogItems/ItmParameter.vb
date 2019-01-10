Imports System
Imports Autodesk
Imports Autodesk.Revit

Public Class ItmParameter

    Private m_Parameter As DB.Parameter

    '仕上編集のために必要な変数
    Public FItem1 As String = ""
    Public FItem2 As String = ""
    Public FItem3 As String = ""
    Public FItem4 As String = ""
    Public FItem5 As String = ""
    Public FItem6 As String = ""


    Public ReadOnly Property Parameter As DB.Parameter
        Get
            Return m_Parameter
        End Get
    End Property


    Public Sub New(ByVal Param As DB.Parameter)
        m_Parameter = Param

    End Sub

    Public ReadOnly Property IsShared As Boolean
        Get
            Return m_Parameter.IsShared
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return Parameter.Definition.Name
    End Function
End Class

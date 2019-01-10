Imports System
Imports Autodesk
Imports Autodesk.Revit

Public Class ItmCategory

    Private m_Category As DB.Category

    Public ReadOnly Property Category As DB.Category
        Get
            Return m_Category
        End Get
    End Property

    Public Sub New(ByVal Cat As DB.Category)
        m_Category = Cat
    End Sub



    Public Overrides Function ToString() As String
        Return m_Category.Name
    End Function
End Class

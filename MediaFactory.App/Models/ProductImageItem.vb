Imports System.Windows.Media.Imaging
Imports MediaFactory.Infrastructure

Namespace Models
    Public Class ProductImageItem
        Inherits ObservableObject

        Private _badgeText As String = String.Empty

        Public Property FullPath As String = String.Empty

        Public Property FileName As String = String.Empty

        Public Property Preview As BitmapImage

        Public Property BadgeText As String
            Get
                Return _badgeText
            End Get
            Set(value As String)
                SetProperty(_badgeText, value)
            End Set
        End Property
    End Class
End Namespace

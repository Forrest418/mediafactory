Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports MediaFactory.Infrastructure

Namespace Models
    Public Class GeneratedImageItem
        Inherits ObservableObject

        Private _sequence As Integer
        Private _title As String = String.Empty
        Private _description As String = String.Empty
        Private _savedPath As String = String.Empty
        Private _preview As BitmapImage
        Private _statusText As String = "Pending"
        Private _statusBrush As Brush = New SolidColorBrush(Color.FromRgb(&H94, &HA3, &HB8))
        Private _lastError As String = String.Empty

        Public Property Sequence As Integer
            Get
                Return _sequence
            End Get
            Set(value As Integer)
                SetProperty(_sequence, value)
            End Set
        End Property

        Public Property Title As String
            Get
                Return _title
            End Get
            Set(value As String)
                SetProperty(_title, value)
            End Set
        End Property

        Public Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                SetProperty(_description, value)
            End Set
        End Property

        Public Property SavedPath As String
            Get
                Return _savedPath
            End Get
            Set(value As String)
                SetProperty(_savedPath, value)
            End Set
        End Property

        Public Property Preview As BitmapImage
            Get
                Return _preview
            End Get
            Set(value As BitmapImage)
                If SetProperty(_preview, value) Then
                    OnPropertyChanged(NameOf(HasPreview))
                End If
            End Set
        End Property

        Public Property StatusText As String
            Get
                Return _statusText
            End Get
            Set(value As String)
                SetProperty(_statusText, value)
            End Set
        End Property

        Public Property StatusBrush As Brush
            Get
                Return _statusBrush
            End Get
            Set(value As Brush)
                SetProperty(_statusBrush, value)
            End Set
        End Property

        Public Property LastError As String
            Get
                Return _lastError
            End Get
            Set(value As String)
                SetProperty(_lastError, value)
            End Set
        End Property

        Public ReadOnly Property HasPreview As Boolean
            Get
                Return Preview IsNot Nothing
            End Get
        End Property
    End Class
End Namespace

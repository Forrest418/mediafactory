Imports System
Imports System.IO
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports MediaFactory.Infrastructure
Imports MediaFactory.ViewModels

Namespace Models
    Public Class MediaLibraryItem
        Inherits ObservableObject

        Private _title As String = String.Empty
        Private _subtitle As String = String.Empty
        Private _projectName As String = String.Empty
        Private _kindText As String = String.Empty
        Private _statusText As String = String.Empty
        Private _filePath As String = String.Empty
        Private _preview As BitmapImage
        Private _statusBrush As Brush
        Private _sortDate As DateTime
        Private _isSourceImage As Boolean
        Private _isGeneratedImage As Boolean
        Private _isFailed As Boolean
        Private _ownerProject As ProjectSessionViewModel

        Public Property Title As String
            Get
                Return _title
            End Get
            Set(value As String)
                SetProperty(_title, value)
            End Set
        End Property

        Public Property Subtitle As String
            Get
                Return _subtitle
            End Get
            Set(value As String)
                SetProperty(_subtitle, value)
            End Set
        End Property

        Public Property ProjectName As String
            Get
                Return _projectName
            End Get
            Set(value As String)
                SetProperty(_projectName, value)
            End Set
        End Property

        Public Property KindText As String
            Get
                Return _kindText
            End Get
            Set(value As String)
                SetProperty(_kindText, value)
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

        Public Property FilePath As String
            Get
                Return _filePath
            End Get
            Set(value As String)
                If SetProperty(_filePath, value) Then
                    OnPropertyChanged(NameOf(FileName))
                    OnPropertyChanged(NameOf(HasFile))
                End If
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

        Public Property StatusBrush As Brush
            Get
                Return _statusBrush
            End Get
            Set(value As Brush)
                SetProperty(_statusBrush, value)
            End Set
        End Property

        Public Property SortDate As DateTime
            Get
                Return _sortDate
            End Get
            Set(value As DateTime)
                SetProperty(_sortDate, value)
            End Set
        End Property

        Public Property IsSourceImage As Boolean
            Get
                Return _isSourceImage
            End Get
            Set(value As Boolean)
                SetProperty(_isSourceImage, value)
            End Set
        End Property

        Public Property IsGeneratedImage As Boolean
            Get
                Return _isGeneratedImage
            End Get
            Set(value As Boolean)
                SetProperty(_isGeneratedImage, value)
            End Set
        End Property

        Public Property IsFailed As Boolean
            Get
                Return _isFailed
            End Get
            Set(value As Boolean)
                SetProperty(_isFailed, value)
            End Set
        End Property

        Public Property OwnerProject As ProjectSessionViewModel
            Get
                Return _ownerProject
            End Get
            Set(value As ProjectSessionViewModel)
                SetProperty(_ownerProject, value)
            End Set
        End Property

        Public ReadOnly Property FileName As String
            Get
                If String.IsNullOrWhiteSpace(FilePath) Then
                    Return String.Empty
                End If

                Return Path.GetFileName(FilePath)
            End Get
        End Property

        Public ReadOnly Property HasPreview As Boolean
            Get
                Return Preview IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property HasFile As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(FilePath) AndAlso File.Exists(FilePath)
            End Get
        End Property
    End Class
End Namespace

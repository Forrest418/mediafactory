Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Linq
Imports MediaFactory.Infrastructure

Namespace Models
    Public Class ModelProviderEditorState
        Public Sub New()
            PlannerProviders = New List(Of ModelProviderEditor)()
            ImageProviders = New List(Of ModelProviderEditor)()
        End Sub

        Public Property PlannerProviders As List(Of ModelProviderEditor)
        Public Property ImageProviders As List(Of ModelProviderEditor)
    End Class

    Public Class ModelProviderEditor
        Inherits ObservableObject

        Private _categoryKey As String = String.Empty
        Private _providerKey As String = String.Empty
        Private _displayName As String = String.Empty
        Private _hint As String = String.Empty

        Public Sub New()
            Fields = New ObservableCollection(Of ModelProviderFieldEditor)()
        End Sub

        Public Property CategoryKey As String
            Get
                Return _categoryKey
            End Get
            Set(value As String)
                SetProperty(_categoryKey, NormalizeText(value))
            End Set
        End Property

        Public Property ProviderKey As String
            Get
                Return _providerKey
            End Get
            Set(value As String)
                SetProperty(_providerKey, NormalizeText(value))
            End Set
        End Property

        Public Property DisplayName As String
            Get
                Return _displayName
            End Get
            Set(value As String)
                SetProperty(_displayName, NormalizeText(value))
            End Set
        End Property

        Public Property Hint As String
            Get
                Return _hint
            End Get
            Set(value As String)
                If SetProperty(_hint, NormalizeText(value)) Then
                    OnPropertyChanged(NameOf(HasHint))
                End If
            End Set
        End Property

        Public ReadOnly Property Fields As ObservableCollection(Of ModelProviderFieldEditor)

        Public ReadOnly Property HasHint As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(Hint)
            End Get
        End Property

        Private Shared Function NormalizeText(value As String) As String
            Return If(value, String.Empty).Trim()
        End Function
    End Class

    Public Class ModelProviderFieldEditor
        Inherits ObservableObject

        Private _fieldKey As String = String.Empty
        Private _displayName As String = String.Empty
        Private _value As String = String.Empty

        Public Sub New()
            Options = New ObservableCollection(Of String)()
        End Sub

        Public Property FieldKey As String
            Get
                Return _fieldKey
            End Get
            Set(value As String)
                SetProperty(_fieldKey, NormalizeText(value))
            End Set
        End Property

        Public Property DisplayName As String
            Get
                Return _displayName
            End Get
            Set(value As String)
                SetProperty(_displayName, NormalizeText(value))
            End Set
        End Property

        Public Property Value As String
            Get
                Return _value
            End Get
            Set(value As String)
                SetProperty(_value, NormalizeText(value))
            End Set
        End Property

        Public ReadOnly Property Options As ObservableCollection(Of String)

        Public ReadOnly Property HasOptions As Boolean
            Get
                Return Options.Count > 0
            End Get
        End Property

        Public Sub ReplaceOptions(values As IEnumerable(Of String))
            Dim nextValues = If(values, Enumerable.Empty(Of String)()).
                Select(Function(item) NormalizeText(item)).
                Where(Function(item) Not String.IsNullOrWhiteSpace(item)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()

            Options.Clear()
            For Each item In nextValues
                Options.Add(item)
            Next

            OnPropertyChanged(NameOf(HasOptions))
        End Sub

        Private Shared Function NormalizeText(value As String) As String
            Return If(value, String.Empty).Trim()
        End Function
    End Class
End Namespace

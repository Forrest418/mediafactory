Imports System
Imports MediaFactory.Infrastructure

Namespace Models
    Public Class ScenarioPreset
        Inherits ObservableObject

        Private _id As String = Guid.NewGuid().ToString("N")
        Private _name As String = String.Empty
        Private _description As String = String.Empty
        Private _designPlanningTemplate As String = String.Empty
        Private _imagePlanningTemplate As String = String.Empty
        Private _isBuiltIn As Boolean
        Private _sortOrder As Integer

        Public Property Id As String
            Get
                Return _id
            End Get
            Set(value As String)
                SetProperty(_id, value)
            End Set
        End Property

        Public Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                Dim normalized = If(value, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalized) Then
                    normalized = "未命名场景"
                End If

                SetProperty(_name, normalized)
            End Set
        End Property

        Public Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                SetProperty(_description, If(value, String.Empty).Trim())
            End Set
        End Property

        Public Property DesignPlanningTemplate As String
            Get
                Return _designPlanningTemplate
            End Get
            Set(value As String)
                SetProperty(_designPlanningTemplate, If(value, String.Empty).Trim())
            End Set
        End Property

        Public Property ImagePlanningTemplate As String
            Get
                Return _imagePlanningTemplate
            End Get
            Set(value As String)
                SetProperty(_imagePlanningTemplate, If(value, String.Empty).Trim())
            End Set
        End Property

        Public Property PlanningInstruction As String
            Get
                Return DesignPlanningTemplate
            End Get
            Set(value As String)
                DesignPlanningTemplate = value
            End Set
        End Property

        Public Property ImageInstruction As String
            Get
                Return ImagePlanningTemplate
            End Get
            Set(value As String)
                ImagePlanningTemplate = value
            End Set
        End Property

        Public Property IsBuiltIn As Boolean
            Get
                Return _isBuiltIn
            End Get
            Set(value As Boolean)
                If SetProperty(_isBuiltIn, value) Then
                    OnPropertyChanged(NameOf(SourceLabel))
                End If
            End Set
        End Property

        Public Property SortOrder As Integer
            Get
                Return _sortOrder
            End Get
            Set(value As Integer)
                SetProperty(_sortOrder, value)
            End Set
        End Property

        Public ReadOnly Property SourceLabel As String
            Get
                Return If(IsBuiltIn, "Built-in", "Custom")
            End Get
        End Property
    End Class
End Namespace

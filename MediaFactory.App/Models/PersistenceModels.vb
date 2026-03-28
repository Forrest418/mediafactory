Imports System.Collections.Generic

Namespace Models
    Public Class PersistedProjectRecord
        Public Property ProjectId As String = String.Empty
        Public Property ProjectName As String = String.Empty
        Public Property IsArchived As Boolean
        Public Property ProjectSummary As String = String.Empty
        Public Property SelectedScenarioPresetId As String = String.Empty
        Public Property SelectedPlannerProviderKey As String = String.Empty
        Public Property RequirementText As String = String.Empty
        Public Property TargetLanguage As String = String.Empty
        Public Property SelectedPlannerModel As String = String.Empty
        Public Property SelectedImageProviderKey As String = String.Empty
        Public Property SelectedImageModel As String = String.Empty
        Public Property SelectedAspectRatio As String = String.Empty
        Public Property OutputResolution As String = String.Empty
        Public Property RequestedCount As Integer
        Public Property DerivativeImageCount As Integer = 1
        Public Property AutoRenderAfterPlan As Boolean
        Public Property DesignPlanMarkdown As String = String.Empty
        Public Property ImagePlanMarkdown As String = String.Empty
        Public Property HasGeneratedPlan As Boolean
        Public Property LastOutputDirectory As String = String.Empty
        Public Property LastActivatedAt As DateTime
        Public Property UploadedImagePaths As List(Of String) = New List(Of String)()
        Public Property GeneratedImages As List(Of PersistedGeneratedImageRecord) = New List(Of PersistedGeneratedImageRecord)()
    End Class

    Public Class PersistedScenarioPresetRecord
        Public Property Id As String = String.Empty
        Public Property Name As String = String.Empty
        Public Property Description As String = String.Empty
        Public Property DesignPlanningTemplate As String = String.Empty
        Public Property ImagePlanningTemplate As String = String.Empty
        Public Property IsBuiltIn As Boolean
        Public Property SortOrder As Integer

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
    End Class

    Public Class PersistedGeneratedImageRecord
        Public Property Sequence As Integer
        Public Property Title As String = String.Empty
        Public Property Description As String = String.Empty
        Public Property SavedPath As String = String.Empty
        Public Property StatusText As String = String.Empty
        Public Property LastError As String = String.Empty
    End Class

    Public Class PersistedStudioState
        Public Property SelectedProjectId As String = String.Empty
        Public Property CurrentPage As Integer
        Public Property SearchText As String = String.Empty
        Public Property ActiveFilter As Integer
        Public Property LanguageCode As String = "en"
    End Class
End Namespace

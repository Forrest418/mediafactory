Namespace Models
    Public Class AnalyzeRequest
        Public Property RequirementText As String = String.Empty
        Public Property TargetLanguage As String = String.Empty
        Public Property PlanningLanguageCode As String = "en"
        Public Property AspectRatio As String = String.Empty
        Public Property RequestedCount As Integer
        Public Property ScenarioName As String = String.Empty
        Public Property ScenarioDescription As String = String.Empty
        Public Property ScenarioDesignPlanningTemplate As String = String.Empty
        Public Property ScenarioImagePlanningTemplate As String = String.Empty

        ' Backward-compatible aliases for incremental migration.
        Public Property ScenarioPlanningInstruction As String
            Get
                Return ScenarioDesignPlanningTemplate
            End Get
            Set(value As String)
                ScenarioDesignPlanningTemplate = value
            End Set
        End Property

        Public Property ScenarioImageInstruction As String
            Get
                Return ScenarioImagePlanningTemplate
            End Get
            Set(value As String)
                ScenarioImagePlanningTemplate = value
            End Set
        End Property
    End Class
End Namespace

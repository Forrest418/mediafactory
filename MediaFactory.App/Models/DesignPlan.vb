Namespace Models
    Public Class DesignPlan
        Public Property ProductSummary As String = String.Empty

        Public Property DesignTheme As String = String.Empty

        Public Property Audience As String = String.Empty

        Public Property ColorSystem As String = String.Empty

        Public Property Typography As String = String.Empty

        Public Property VisualLanguage As String = String.Empty

        Public Property PhotographyStyle As String = String.Empty

        Public Property LayoutGuidance As String = String.Empty

        Public Property ImagePlans As List(Of ImagePlanItem) = New List(Of ImagePlanItem)()
    End Class
End Namespace

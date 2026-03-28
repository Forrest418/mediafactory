Imports System.Collections.ObjectModel

Namespace Models
    Public Class ProjectSummaryItem
        Public Property ProjectName As String = String.Empty
        Public Property ProductName As String = String.Empty
        Public Property Market As String = String.Empty
        Public Property Stage As String = String.Empty
        Public Property Progress As Integer
        Public Property Status As String = String.Empty
        Public Property UpdatedAt As String = String.Empty
        Public Property Owner As String = String.Empty
    End Class

    Public Class ProjectAssetItem
        Public Property Name As String = String.Empty
        Public Property Category As String = String.Empty
        Public Property Tag As String = String.Empty
        Public Property Description As String = String.Empty
    End Class

    Public Class ProjectPlanItem
        Public Property Sequence As Integer
        Public Property Title As String = String.Empty
        Public Property Scene As String = String.Empty
        Public Property SellingPoint As String = String.Empty
        Public Property LayoutHint As String = String.Empty
        Public Property Status As String = String.Empty
    End Class

    Public Class ProjectOutputItem
        Public Property Title As String = String.Empty
        Public Property VariantName As String = String.Empty
        Public Property Channel As String = String.Empty
        Public Property Status As String = String.Empty
        Public Property Notes As String = String.Empty
    End Class

    Public Class ProjectTaskItem
        Public Property ProjectName As String = String.Empty
        Public Property Stage As String = String.Empty
        Public Property Status As String = String.Empty
        Public Property Progress As Integer
        Public Property Worker As String = String.Empty
        Public Property UpdatedAt As String = String.Empty
    End Class

    Public Class ProjectWorkspace
        Public Property DisplayName As String = String.Empty
        Public Property Subtitle As String = String.Empty
        Public Property ProductName As String = String.Empty
        Public Property Brand As String = String.Empty
        Public Property Market As String = String.Empty
        Public Property Language As String = String.Empty
        Public Property AspectRatio As String = String.Empty
        Public Property ModelName As String = String.Empty
        Public Property StageName As String = String.Empty
        Public Property StatusText As String = String.Empty
        Public Property CreativeDirection As String = String.Empty
        Public Property BrandNotes As String = String.Empty
        Public Property QueueSummary As String = String.Empty
        Public Property LastUpdate As String = String.Empty
        Public Property LeadDesigner As String = String.Empty
        Public Property TargetAudience As String = String.Empty
        Public Property TargetChannels As String = String.Empty
        Public Property PlannedCount As Integer
        Public Property OutputResolution As String = String.Empty
        Public Property ReviewOwner As String = String.Empty
        Public Property Assets As ObservableCollection(Of ProjectAssetItem) = New ObservableCollection(Of ProjectAssetItem)()
        Public Property Plans As ObservableCollection(Of ProjectPlanItem) = New ObservableCollection(Of ProjectPlanItem)()
        Public Property Outputs As ObservableCollection(Of ProjectOutputItem) = New ObservableCollection(Of ProjectOutputItem)()
    End Class
End Namespace

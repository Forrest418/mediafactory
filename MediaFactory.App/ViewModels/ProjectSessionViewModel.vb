Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Threading
Imports MediaFactory.Infrastructure
Imports MediaFactory.Localization
Imports MediaFactory.Models
Imports MediaFactory.Services
Imports MediaFactory.Utils

Namespace ViewModels
    Public Class ProjectSessionViewModel
        Inherits ObservableObject
        Implements IDisposable

        Private Shared ReadOnly PendingBrush As Brush = CreateBrush("#94A3B8")
        Private Shared ReadOnly RunningBrush As Brush = CreateBrush("#0EA5E9")
        Private Shared ReadOnly SuccessBrush As Brush = CreateBrush("#22C55E")
        Private Shared ReadOnly FailedBrush As Brush = CreateBrush("#EF4444")
        Private Shared ReadOnly MissingBrush As Brush = CreateBrush("#F59E0B")
        Private Shared ReadOnly PrimaryActionBrush As Brush = CreateBrush("#2563EB")
        Private Shared ReadOnly CompletedActionBrush As Brush = CreateBrush("#1D4ED8")
        Private Shared ReadOnly DisabledActionBrush As Brush = CreateBrush("#CBD5E1")
        Private Shared ReadOnly StopActionBrush As Brush = CreateBrush("#DC2626")
        Private Const MaxImageGenerationAttempts As Integer = 3
        Private Const ImageRetryDelayMinMilliseconds As Integer = 3000
        Private Const ImageRetryDelayMaxMilliseconds As Integer = 5000
        Private Const DefaultProjectSummaryText As String = "No project summary yet"
        Private Const AutoProjectSummaryMaxLength As Integer = 48

        Private Const NoTextLanguage As String = "No Text"
        Public Const DefaultExecutionSummary As String = "Import images -> generate editable plan markdown -> review -> generate images."

        Private ReadOnly _vertexService As VertexAiService
        Private ReadOnly _logger As AppLogger
        Private ReadOnly _dispatcher As Dispatcher
        Private ReadOnly _refreshTimer As DispatcherTimer

        Private _outputWatcher As FileSystemWatcher
        Private _operationCancellationTokenSource As CancellationTokenSource
        Private _projectId As String
        Private _projectName As String
        Private _isArchived As Boolean
        Private _projectSummary As String = DefaultProjectSummaryText
        Private _availableScenarioPresets As ObservableCollection(Of ScenarioPreset)
        Private _selectedScenarioPresetId As String = String.Empty
        Private _statusMessage As String = "Ready"
        Private _executionSummary As String = DefaultExecutionSummary
        Private _requirementText As String = "This is an outdoor gas grill for the North American market. Highlight durability, cooking area, and backyard usage."
        Private _targetLanguage As String = NoTextLanguage
        Private _selectedPlannerProviderKey As String = String.Empty
        Private _selectedPlannerModel As String = "Auto"
        Private _selectedImageProviderKey As String = String.Empty
        Private _selectedImageModel As String = String.Empty
        Private _selectedAspectRatio As String = "1:1"
        Private _outputResolution As String = "1K"
        Private _requestedCount As Integer = 4
        Private _derivativeImageCount As Integer = 1
        Private _designPlanMarkdown As String = String.Empty
        Private _imagePlanMarkdown As String = String.Empty
        Private _isBusy As Boolean
        Private _isPlanning As Boolean
        Private _isGeneratingImages As Boolean
        Private _hasGeneratedPlan As Boolean
        Private _lastOutputDirectory As String = String.Empty
        Private _selectedGeneratedImage As GeneratedImageItem
        Private _recentLogText As String = String.Empty
        Private _isPinned As Boolean
        Private _lastActivatedAt As DateTime = DateTime.Now
        Private _autoRenderAfterPlan As Boolean
        Private _isQueueSelected As Boolean

        Public Sub New(projectName As String, vertexService As VertexAiService, Optional projectId As String = Nothing)
            _projectId = If(String.IsNullOrWhiteSpace(projectId), Guid.NewGuid().ToString("N"), projectId)
            _projectName = projectName
            _vertexService = vertexService
            _logger = New AppLogger()
            _dispatcher = Application.Current.Dispatcher
            _refreshTimer = New DispatcherTimer With {.Interval = TimeSpan.FromMilliseconds(450)}
            AddHandler _refreshTimer.Tick, AddressOf RefreshTimer_Tick

            UploadedImages = New ObservableCollection(Of ProductImageItem)()
            GeneratedImages = New ObservableCollection(Of GeneratedImageItem)()
            _availableScenarioPresets = New ObservableCollection(Of ScenarioPreset)()
            AddHandler GeneratedImages.CollectionChanged, AddressOf GeneratedImages_CollectionChanged

            RefreshProviderConfiguration()
            EnsureProjectSummaryFromRequirement(True)
            WriteLog($"Project created: {projectName}")
            WriteLog($"Log file: {_logger.LogFilePath}")
        End Sub

        Public ReadOnly Property ProjectId As String
            Get
                Return _projectId
            End Get
        End Property

        Public ReadOnly Property UploadedImages As ObservableCollection(Of ProductImageItem)
        Public ReadOnly Property GeneratedImages As ObservableCollection(Of GeneratedImageItem)

        Public ReadOnly Property AvailableLanguages As IReadOnlyList(Of String) = New List(Of String) From {
            NoTextLanguage, "English", "Chinese", "Japanese", "Korean", "German", "French", "Spanish", "Italian", "Portuguese", "Russian", "Arabic"
        }

        Public ReadOnly Property AvailableAspectRatios As IReadOnlyList(Of String) = New List(Of String) From {
            "1:1", "2:3", "3:2", "3:4", "4:3", "4:5", "5:4", "9:16", "16:9", "21:9"
        }

        Public ReadOnly Property AvailablePlannerProviderOptions As IReadOnlyList(Of ProviderOptionItem) = New List(Of ProviderOptionItem) From {
            New ProviderOptionItem With {.Key = "deepseek", .DisplayName = "DeepSeek"},
            New ProviderOptionItem With {.Key = "openai", .DisplayName = "OpenAI Compatible"},
            New ProviderOptionItem With {.Key = "qwen", .DisplayName = "Qwen"}
        }

        Public ReadOnly Property AvailableImageProviderOptions As IReadOnlyList(Of ProviderOptionItem) = New List(Of ProviderOptionItem) From {
            New ProviderOptionItem With {.Key = "gemini", .DisplayName = "Gemini"},
            New ProviderOptionItem With {.Key = "openai", .DisplayName = "OpenAI Compatible"},
            New ProviderOptionItem With {.Key = "qwen", .DisplayName = "Qwen"}
        }

        Public ReadOnly Property AvailableImageSizes As IReadOnlyList(Of String) = New List(Of String) From {"1K", "2K", "4K"}
        Public ReadOnly Property AvailableGenerationCounts As IReadOnlyList(Of Integer) = Enumerable.Range(1, 10).ToList()
        Public ReadOnly Property AvailableDerivativeCounts As IReadOnlyList(Of Integer) = Enumerable.Range(1, 10).ToList()
        Public ReadOnly Property AvailableScenarioPresets As ObservableCollection(Of ScenarioPreset)
            Get
                Return _availableScenarioPresets
            End Get
        End Property

        Public Property ProjectName As String
            Get
                Return _projectName
            End Get
            Set(value As String)
                If SetProperty(_projectName, value) Then
                    OnPropertyChanged(NameOf(ProjectTabCaption))
                End If
            End Set
        End Property

        Public Property IsArchived As Boolean
            Get
                Return _isArchived
            End Get
            Set(value As Boolean)
                SetProperty(_isArchived, value)
            End Set
        End Property

        Public Property ProjectSummary As String
            Get
                Return _projectSummary
            End Get
            Set(value As String)
                Dim normalized = If(value, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalized) Then
                    normalized = DefaultProjectSummaryText
                End If

                If SetProperty(_projectSummary, normalized) Then
                    OnPropertyChanged(NameOf(ProjectSubtitle))
                End If
            End Set
        End Property

        Public Property SelectedScenarioPresetId As String
            Get
                Return _selectedScenarioPresetId
            End Get
            Set(value As String)
                Dim normalized = If(value, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalized) AndAlso _availableScenarioPresets.Count > 0 Then
                    normalized = _availableScenarioPresets(0).Id
                End If

                If SetProperty(_selectedScenarioPresetId, normalized) Then
                    OnPropertyChanged(NameOf(SelectedScenarioPresetName))
                    OnPropertyChanged(NameOf(SelectedScenarioPresetDescription))
                    OnPropertyChanged(NameOf(ScenarioSummaryText))
                End If
            End Set
        End Property

        Public ReadOnly Property SelectedScenarioPresetName As String
            Get
                Return If(GetSelectedScenarioPreset()?.Name, "未选择场景")
            End Get
        End Property

        Public ReadOnly Property SelectedScenarioPresetDescription As String
            Get
                Return If(GetSelectedScenarioPreset()?.Description, String.Empty)
            End Get
        End Property

        Public ReadOnly Property ProjectTabCaption As String
            Get
                Return ProjectName
            End Get
        End Property

        Public ReadOnly Property LocalizedStageText As String
            Get
                If IsPlanning Then Return T("Project.Stage.Planning", "Planning")
                If IsGeneratingImages Then Return T("Project.Stage.Rendering", "Rendering")
                If HasGeneratedPlan Then Return T("Project.Stage.ReadyToRender", "Ready To Render")
                If UploadedImages.Count > 0 Then Return T("Project.Stage.ReadyToPlan", "Ready To Plan")
                Return T("Project.Stage.Idle", "Idle")
            End Get
        End Property

        Public ReadOnly Property ProjectStageText As String
            Get
                If IsPlanning Then Return "Planning"
                If IsGeneratingImages Then Return "Generating"
                If HasGeneratedPlan Then Return "Ready To Render"
                If UploadedImages.Count > 0 Then Return "Ready To Plan"
                Return "Idle"
            End Get
        End Property

        Public ReadOnly Property StageBrush As Brush
            Get
                If IsPlanning OrElse IsGeneratingImages Then Return RunningBrush
                If GeneratedImages.Any(Function(item) item.StatusText = "Success") Then Return SuccessBrush
                If HasGeneratedPlan Then Return CompletedActionBrush
                If UploadedImages.Count > 0 Then Return PrimaryActionBrush
                Return PendingBrush
            End Get
        End Property

        Public ReadOnly Property ProgressPercent As Integer
            Get
                If IsGeneratingImages Then
                    Dim total = Math.Max(1, EffectiveRenderItemCount)
                    Dim finished = GeneratedImages.Where(Function(item) item.StatusText = "Success" OrElse item.StatusText = "Failed").Count()
                    Return Math.Min(100, CInt(Math.Round(finished * 100.0 / total)))
                End If

                If IsPlanning Then Return 45
                If HasGeneratedPlan Then Return 72
                If UploadedImages.Count > 0 Then Return 25
                Return 0
            End Get
        End Property

        Public ReadOnly Property ProgressLabelText As String
            Get
                Return $"{ProgressPercent}%"
            End Get
        End Property

        Public ReadOnly Property ProjectSubtitle As String
            Get
                Return ProjectSummary
            End Get
        End Property

        Public ReadOnly Property AssetCountText As String
            Get
                Dim assetCount = Math.Max(GeneratedImages.Count, UploadedImages.Count)
                Return F("Project.AssetCount", "{0} assets", assetCount)
            End Get
        End Property

        Public ReadOnly Property SourceSummaryText As String
            Get
                Return F("Project.SourceSummary", "{0} source / {1} planned / {2} outputs", UploadedImages.Count, RequestedCount, EffectiveRenderItemCount)
            End Get
        End Property

        Public ReadOnly Property EffectiveRenderItemCount As Integer
            Get
                Return RequestedCount * DerivativeImageCount
            End Get
        End Property

        Public ReadOnly Property ScenarioSummaryText As String
            Get
                Return SelectedScenarioPresetName
            End Get
        End Property

        Public ReadOnly Property SuccessfulImageCount As Integer
            Get
                Return GeneratedImages.Where(Function(item) String.Equals(item.StatusText, "Success", StringComparison.Ordinal)).Count()
            End Get
        End Property

        Public ReadOnly Property RetryableFailureCount As Integer
            Get
                Return GeneratedImages.Where(Function(item) String.Equals(item.StatusText, "Failed", StringComparison.Ordinal) OrElse
                                                     String.Equals(item.StatusText, "Missing", StringComparison.Ordinal)).Count()
            End Get
        End Property

        Public ReadOnly Property HasRetryableFailures As Boolean
            Get
                Return RetryableFailureCount > 0
            End Get
        End Property

        Public ReadOnly Property FailureSummaryText As String
            Get
                If RetryableFailureCount <= 0 Then
                    Return T("Project.Failure.None", "No issues")
                End If

                Return $"{RetryableFailureCount} failed"
            End Get
        End Property

        Public ReadOnly Property FailureVisibility As Visibility
            Get
                Return If(HasRetryableFailures, Visibility.Visible, Visibility.Collapsed)
            End Get
        End Property

        Public ReadOnly Property LastUpdatedText As String
            Get
                Dim elapsed = DateTime.Now - LastActivatedAt
                If elapsed.TotalMinutes < 1 Then Return T("Project.Updated.JustNow", "Updated just now")
                If elapsed.TotalHours < 1 Then Return F("Project.Updated.Minutes", "{0} min ago", Math.Max(1, CInt(Math.Floor(elapsed.TotalMinutes))))
                If elapsed.TotalDays < 1 Then Return F("Project.Updated.Hours", "{0} hr ago", Math.Max(1, CInt(Math.Floor(elapsed.TotalHours))))
                Return F("Project.Updated.Days", "{0} day(s) ago", Math.Max(1, CInt(Math.Floor(elapsed.TotalDays))))
            End Get
        End Property

        Public ReadOnly Property QueueRemainingText As String
            Get
                If IsPlanning Then Return T("Project.Queue.Planning", "Planning")
                If IsGeneratingImages Then Return T("Project.Queue.Processing", "Processing")
                If HasRetryableFailures Then Return T("Project.Queue.RetryFailed", "Failed, retry needed")
                If GeneratedImages.Any(Function(item) String.Equals(item.StatusText, "Success", StringComparison.Ordinal)) Then Return T("Project.Queue.Completed", "Completed")
                If UploadedImages.Count > 0 Then Return T("Project.Queue.Waiting", "Waiting")
                Return T("Project.Queue.Empty", "--")
            End Get
        End Property

        Public ReadOnly Property QueueActionText As String
            Get
                If IsPlanning OrElse IsGeneratingImages Then Return T("Project.Action.Stop", "Stop")
                If HasRetryableFailures Then Return T("Project.Action.RetryFailed", "Retry Failed")
                If UploadedImages.Count > 0 Then
                    If HasGeneratedPlan Then
                        Return T("Project.Action.StartRender", "Start Rendering")
                    End If

                    Return T("Project.Action.StartPlan", "Start Planning")
                End If

                Return T("Project.Action.Open", "Open")
            End Get
        End Property

        Public ReadOnly Property QueueActionBackground As Brush
            Get
                If IsPlanning OrElse IsGeneratingImages Then Return StopActionBrush
                If HasRetryableFailures Then Return Brushes.White
                Return PrimaryActionBrush
            End Get
        End Property

        Public ReadOnly Property QueueActionBorderBrush As Brush
            Get
                If IsPlanning OrElse IsGeneratingImages Then Return StopActionBrush
                If HasRetryableFailures Then Return FailedBrush
                Return PrimaryActionBrush
            End Get
        End Property

        Public ReadOnly Property QueueActionForeground As Brush
            Get
                If HasRetryableFailures AndAlso Not IsPlanning AndAlso Not IsGeneratingImages Then
                    Return FailedBrush
                End If

                Return Brushes.White
            End Get
        End Property

        Public ReadOnly Property ProgressBrush As Brush
            Get
                If HasRetryableFailures Then Return FailedBrush
                If IsGeneratingImages Then Return RunningBrush
                If SuccessfulImageCount > 0 Then Return SuccessBrush
                If HasGeneratedPlan Then Return CompletedActionBrush
                If UploadedImages.Count > 0 Then Return PrimaryActionBrush
                Return PendingBrush
            End Get
        End Property

        Public ReadOnly Property ProjectCover As BitmapImage
            Get
                Dim generatedPreview = GeneratedImages.
                    OrderBy(Function(item) item.Sequence).
                    Select(Function(item) item.Preview).
                    FirstOrDefault(Function(preview) preview IsNot Nothing)
                If generatedPreview IsNot Nothing Then
                    Return generatedPreview
                End If

                Return UploadedImages.Select(Function(item) item.Preview).FirstOrDefault(Function(preview) preview IsNot Nothing)
            End Get
        End Property

        Public ReadOnly Property HasProjectCover As Boolean
            Get
                Return ProjectCover IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property StepInputBrush As Brush
            Get
                Return If(UploadedImages.Count > 0, PrimaryActionBrush, PendingBrush)
            End Get
        End Property

        Public ReadOnly Property StepPlanBrush As Brush
            Get
                If IsPlanning OrElse HasGeneratedPlan OrElse IsGeneratingImages Then
                    Return CompletedActionBrush
                End If

                Return PendingBrush
            End Get
        End Property

        Public ReadOnly Property StepRenderBrush As Brush
            Get
                If IsGeneratingImages Then Return RunningBrush
                If GeneratedImages.Count > 0 Then Return PrimaryActionBrush
                Return PendingBrush
            End Get
        End Property

        Public ReadOnly Property StepReviewBrush As Brush
            Get
                If GeneratedImages.Any(Function(item) String.Equals(item.StatusText, "Success", StringComparison.Ordinal)) Then
                    Return SuccessBrush
                End If

                Return PendingBrush
            End Get
        End Property

        Public Property IsPinned As Boolean
            Get
                Return _isPinned
            End Get
            Set(value As Boolean)
                SetProperty(_isPinned, value)
            End Set
        End Property

        Public Property LastActivatedAt As DateTime
            Get
                Return _lastActivatedAt
            End Get
            Set(value As DateTime)
                If SetProperty(_lastActivatedAt, value) Then
                    OnPropertyChanged(NameOf(LastUpdatedText))
                End If
            End Set
        End Property

        Public Property RecentLogText As String
            Get
                Return _recentLogText
            End Get
            Set(value As String)
                SetProperty(_recentLogText, value)
            End Set
        End Property

        Public ReadOnly Property LogFilePath As String
            Get
                Return _logger.LogFilePath
            End Get
        End Property

        Public Property StatusMessage As String
            Get
                Return _statusMessage
            End Get
            Set(value As String)
                SetProperty(_statusMessage, value)
            End Set
        End Property

        Public Property ExecutionSummary As String
            Get
                Return _executionSummary
            End Get
            Set(value As String)
                SetProperty(_executionSummary, value)
            End Set
        End Property

        Public Property RequirementText As String
            Get
                Return _requirementText
            End Get
            Set(value As String)
                Dim previousRequirement = _requirementText
                Dim previousFallbackSummary = BuildFallbackProjectSummary(previousRequirement)
                If SetProperty(_requirementText, value) Then
                    If IsDefaultProjectSummary(_projectSummary) OrElse
                       String.Equals(_projectSummary, previousFallbackSummary, StringComparison.Ordinal) Then
                        EnsureProjectSummaryFromRequirement(True)
                    End If
                End If
            End Set
        End Property

        Public Property AutoRenderAfterPlan As Boolean
            Get
                Return _autoRenderAfterPlan
            End Get
            Set(value As Boolean)
                SetProperty(_autoRenderAfterPlan, value)
            End Set
        End Property

        Public Property IsQueueSelected As Boolean
            Get
                Return _isQueueSelected
            End Get
            Set(value As Boolean)
                SetProperty(_isQueueSelected, value)
            End Set
        End Property

        Public Sub BindScenarioPresets(presets As ObservableCollection(Of ScenarioPreset))
            _availableScenarioPresets = If(presets, New ObservableCollection(Of ScenarioPreset)())

            If String.IsNullOrWhiteSpace(_selectedScenarioPresetId) OrElse GetSelectedScenarioPreset() Is Nothing Then
                _selectedScenarioPresetId = _availableScenarioPresets.FirstOrDefault()?.Id
            End If

            OnPropertyChanged(NameOf(AvailableScenarioPresets))
            OnPropertyChanged(NameOf(SelectedScenarioPresetId))
            OnPropertyChanged(NameOf(SelectedScenarioPresetName))
            OnPropertyChanged(NameOf(SelectedScenarioPresetDescription))
            OnPropertyChanged(NameOf(ScenarioSummaryText))
        End Sub

        Public Sub NotifyScenarioPresetsChanged()
            If GetSelectedScenarioPreset() Is Nothing AndAlso _availableScenarioPresets.Count > 0 Then
                _selectedScenarioPresetId = _availableScenarioPresets(0).Id
            End If

            OnPropertyChanged(NameOf(SelectedScenarioPresetName))
            OnPropertyChanged(NameOf(SelectedScenarioPresetDescription))
            OnPropertyChanged(NameOf(ScenarioSummaryText))
        End Sub

        Public Sub RefreshProviderConfiguration()
            EnsureProjectModelSelections()

            If IsCustomImageSizeMode Then
                If String.IsNullOrWhiteSpace(_outputResolution) OrElse Not _outputResolution.Contains("*"c) Then
                    _outputResolution = "1024*1024"
                    OnPropertyChanged(NameOf(OutputResolution))
                End If
            Else
                Dim normalized = If(_outputResolution, String.Empty).Trim().ToUpperInvariant()
                If String.IsNullOrWhiteSpace(normalized) OrElse normalized.Contains("*"c) Then
                    _outputResolution = "1K"
                    OnPropertyChanged(NameOf(OutputResolution))
                End If
            End If

            StatusMessage = F("Project.SelectedModels", "Planner: {0} | Image: {1}", SelectedPlannerDisplayName, SelectedImageDisplayName)
            OnPropertyChanged(NameOf(SelectedPlannerProviderKey))
            OnPropertyChanged(NameOf(SelectedPlannerModel))
            OnPropertyChanged(NameOf(SelectedPlannerDisplayName))
            OnPropertyChanged(NameOf(AvailablePlannerModelOptions))
            OnPropertyChanged(NameOf(SelectedImageProviderKey))
            OnPropertyChanged(NameOf(SelectedImageModel))
            OnPropertyChanged(NameOf(SelectedImageDisplayName))
            OnPropertyChanged(NameOf(AvailableImageModelOptions))
            OnPropertyChanged(NameOf(CurrentImageModelDisplay))
            OnPropertyChanged(NameOf(IsCustomImageSizeMode))
            OnPropertyChanged(NameOf(AspectRatioVisibility))
            OnPropertyChanged(NameOf(ImageSizeComboVisibility))
            OnPropertyChanged(NameOf(ImageSizeTextVisibility))
            OnPropertyChanged(NameOf(ImageSizeSplitVisibility))
            OnPropertyChanged(NameOf(ResolutionLabelText))
            OnPropertyChanged(NameOf(ResolutionHintText))
            OnPropertyChanged(NameOf(CustomImageWidth))
            OnPropertyChanged(NameOf(CustomImageHeight))
        End Sub

        Public Property TargetLanguage As String
            Get
                Return _targetLanguage
            End Get
            Set(value As String)
                SetProperty(_targetLanguage, value)
            End Set
        End Property

        Public Property SelectedPlannerProviderKey As String
            Get
                Return _selectedPlannerProviderKey
            End Get
            Set(value As String)
                Dim normalized = If(value, String.Empty).Trim().ToLowerInvariant()
                If String.IsNullOrWhiteSpace(normalized) Then
                    normalized = _vertexService.GetDefaultPlannerProviderKey()
                End If

                If SetProperty(_selectedPlannerProviderKey, normalized) Then
                    EnsureProjectModelSelections()
                    OnPropertyChanged(NameOf(SelectedPlannerModel))
                    OnPropertyChanged(NameOf(AvailablePlannerModelOptions))
                    OnPropertyChanged(NameOf(SelectedPlannerDisplayName))
                End If
            End Set
        End Property

        Public Property SelectedPlannerModel As String
            Get
                Return _selectedPlannerModel
            End Get
            Set(value As String)
                Dim normalized = If(value, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalized) Then
                    normalized = _vertexService.GetDefaultPlannerModel(SelectedPlannerProviderKey)
                End If

                If SetProperty(_selectedPlannerModel, normalized) Then
                    OnPropertyChanged(NameOf(SelectedPlannerDisplayName))
                End If
            End Set
        End Property

        Public Property SelectedImageProviderKey As String
            Get
                Return _selectedImageProviderKey
            End Get
            Set(value As String)
                Dim normalized = If(value, String.Empty).Trim().ToLowerInvariant()
                If String.IsNullOrWhiteSpace(normalized) Then
                    normalized = _vertexService.GetDefaultImageProviderKey()
                End If

                If SetProperty(_selectedImageProviderKey, normalized) Then
                    EnsureProjectModelSelections()
                    OnPropertyChanged(NameOf(SelectedImageModel))
                    OnPropertyChanged(NameOf(SelectedImageDisplayName))
                    OnPropertyChanged(NameOf(CurrentImageModelDisplay))
                    OnPropertyChanged(NameOf(AvailableImageModelOptions))
                    OnPropertyChanged(NameOf(IsCustomImageSizeMode))
                    OnPropertyChanged(NameOf(AspectRatioVisibility))
                    OnPropertyChanged(NameOf(ImageSizeComboVisibility))
                    OnPropertyChanged(NameOf(ImageSizeTextVisibility))
                    OnPropertyChanged(NameOf(ImageSizeSplitVisibility))
                    OnPropertyChanged(NameOf(ResolutionLabelText))
                    OnPropertyChanged(NameOf(ResolutionHintText))
                End If
            End Set
        End Property

        Public Property SelectedImageModel As String
            Get
                Return _selectedImageModel
            End Get
            Set(value As String)
                Dim normalized = If(value, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalized) Then
                    normalized = _vertexService.GetDefaultImageModel(SelectedImageProviderKey)
                End If

                If SetProperty(_selectedImageModel, normalized) Then
                    OnPropertyChanged(NameOf(SelectedImageDisplayName))
                    OnPropertyChanged(NameOf(CurrentImageModelDisplay))
                End If
            End Set
        End Property

        Public ReadOnly Property AvailablePlannerModelOptions As IReadOnlyList(Of String)
            Get
                Return GetPlannerModelOptionsForProvider(SelectedPlannerProviderKey)
            End Get
        End Property

        Public ReadOnly Property AvailableImageModelOptions As IReadOnlyList(Of String)
            Get
                Return GetImageModelOptionsForProvider(SelectedImageProviderKey)
            End Get
        End Property

        Public ReadOnly Property SelectedPlannerDisplayName As String
            Get
                Return _vertexService.GetPlannerDisplayName(SelectedPlannerProviderKey, SelectedPlannerModel)
            End Get
        End Property

        Public ReadOnly Property SelectedImageDisplayName As String
            Get
                Return _vertexService.GetImageDisplayName(SelectedImageProviderKey, SelectedImageModel)
            End Get
        End Property

        Public ReadOnly Property CurrentImageModelDisplay As String
            Get
                Return SelectedImageDisplayName
            End Get
        End Property

        Public ReadOnly Property IsCustomImageSizeMode As Boolean
            Get
                Return _vertexService.UsesCustomImageSizeForProvider(SelectedImageProviderKey)
            End Get
        End Property

        Public ReadOnly Property AspectRatioVisibility As Visibility
            Get
                Return If(IsCustomImageSizeMode, Visibility.Collapsed, Visibility.Visible)
            End Get
        End Property

        Public ReadOnly Property ImageSizeComboVisibility As Visibility
            Get
                Return If(IsCustomImageSizeMode, Visibility.Collapsed, Visibility.Visible)
            End Get
        End Property

        Public ReadOnly Property ImageSizeTextVisibility As Visibility
            Get
                Return If(IsCustomImageSizeMode, Visibility.Visible, Visibility.Collapsed)
            End Get
        End Property

        Public ReadOnly Property ImageSizeSplitVisibility As Visibility
            Get
                Return If(IsCustomImageSizeMode, Visibility.Visible, Visibility.Collapsed)
            End Get
        End Property

        Public ReadOnly Property ResolutionLabelText As String
            Get
                Return If(IsCustomImageSizeMode,
                          T("Project.Resolution.OutputSize", "Output Size"),
                          T("Project.Resolution.Quality", "Image Size"))
            End Get
        End Property

        Public ReadOnly Property ResolutionHintText As String
            Get
                If IsCustomImageSizeMode Then
                    Return T("Project.CustomSizeHint", "Example: 1536*1024. Total pixels must stay between 512*512 and 2048*2048.")
                End If

                Return T("Project.FixedSizeHint", "Gemini preset sizes: 1K, 2K, 4K.")
            End Get
        End Property

        Public Property SelectedAspectRatio As String
            Get
                Return _selectedAspectRatio
            End Get
            Set(value As String)
                SetProperty(_selectedAspectRatio, value)
            End Set
        End Property

        Public Property OutputResolution As String
            Get
                Return _outputResolution
            End Get
            Set(value As String)
                Dim normalized = If(value, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalized) Then
                    normalized = If(IsCustomImageSizeMode, "1024*1024", "1K")
                End If

                If SetProperty(_outputResolution, normalized) Then
                    OnPropertyChanged(NameOf(CustomImageWidth))
                    OnPropertyChanged(NameOf(CustomImageHeight))
                End If
            End Set
        End Property

        Public Property CustomImageWidth As Integer
            Get
                Return GetParsedCustomImageSize().Item1
            End Get
            Set(value As Integer)
                Dim currentSize = GetParsedCustomImageSize()
                Dim currentHeight = currentSize.Item2
                UpdateCustomImageSize(value, currentHeight)
            End Set
        End Property

        Public Property CustomImageHeight As Integer
            Get
                Return GetParsedCustomImageSize().Item2
            End Get
            Set(value As Integer)
                Dim currentSize = GetParsedCustomImageSize()
                Dim currentWidth = currentSize.Item1
                UpdateCustomImageSize(currentWidth, value)
            End Set
        End Property

        Public Property RequestedCount As Integer
            Get
                Return _requestedCount
            End Get
            Set(value As Integer)
                Dim safeValue = Math.Min(10, Math.Max(1, value))
                If SetProperty(_requestedCount, safeValue) Then
                    OnPropertyChanged(NameOf(SourceSummaryText))
                    OnPropertyChanged(NameOf(EffectiveRenderItemCount))
                    OnPropertyChanged(NameOf(ProgressPercent))
                    OnPropertyChanged(NameOf(ProgressLabelText))
                End If
            End Set
        End Property

        Public Property DerivativeImageCount As Integer
            Get
                Return _derivativeImageCount
            End Get
            Set(value As Integer)
                Dim safeValue = Math.Min(10, Math.Max(1, value))
                If SetProperty(_derivativeImageCount, safeValue) Then
                    OnPropertyChanged(NameOf(SourceSummaryText))
                    OnPropertyChanged(NameOf(EffectiveRenderItemCount))
                    OnPropertyChanged(NameOf(ProgressPercent))
                    OnPropertyChanged(NameOf(ProgressLabelText))
                End If
            End Set
        End Property

        Public Property DesignPlanMarkdown As String
            Get
                Return _designPlanMarkdown
            End Get
            Set(value As String)
                SetProperty(_designPlanMarkdown, value)
            End Set
        End Property

        Public Property ImagePlanMarkdown As String
            Get
                Return _imagePlanMarkdown
            End Get
            Set(value As String)
                SetProperty(_imagePlanMarkdown, value)
            End Set
        End Property

        Public Property IsBusy As Boolean
            Get
                Return _isBusy
            End Get
            Set(value As Boolean)
                If SetProperty(_isBusy, value) Then
                    OnPropertyChanged(NameOf(BusyVisibility))
                    OnPropertyChanged(NameOf(CanGeneratePlan))
                    OnPropertyChanged(NameOf(CanGenerateImages))
                    OnPropertyChanged(NameOf(CanRegenerateSelectedImage))
                    OnPropertyChanged(NameOf(CanPreviewSelectedImage))
                    OnPropertyChanged(NameOf(CanEditSelectedImage))
                    OnPropertyChanged(NameOf(CanDownloadSelectedImage))
                    OnPropertyChanged(NameOf(PlanButtonBackground))
                    OnPropertyChanged(NameOf(ImageButtonBackground))
                    OnPropertyChanged(NameOf(ProjectStageText))
                    OnPropertyChanged(NameOf(LocalizedStageText))
                    OnPropertyChanged(NameOf(StageBrush))
                    OnPropertyChanged(NameOf(ProgressPercent))
                    OnPropertyChanged(NameOf(ProgressLabelText))
                    NotifyProjectStateBadgesChanged()
                    NotifyStepBrushesChanged()
                End If
            End Set
        End Property

        Public Property IsPlanning As Boolean
            Get
                Return _isPlanning
            End Get
            Set(value As Boolean)
                If SetProperty(_isPlanning, value) Then
                    OnPropertyChanged(NameOf(CanGeneratePlan))
                    OnPropertyChanged(NameOf(PlanButtonText))
                    OnPropertyChanged(NameOf(PlanButtonBackground))
                    OnPropertyChanged(NameOf(ProjectStageText))
                    OnPropertyChanged(NameOf(LocalizedStageText))
                    OnPropertyChanged(NameOf(StageBrush))
                    OnPropertyChanged(NameOf(ProgressPercent))
                    OnPropertyChanged(NameOf(ProgressLabelText))
                    NotifyProjectStateBadgesChanged()
                    NotifyStepBrushesChanged()
                End If
            End Set
        End Property

        Public Property IsGeneratingImages As Boolean
            Get
                Return _isGeneratingImages
            End Get
            Set(value As Boolean)
                If SetProperty(_isGeneratingImages, value) Then
                    OnPropertyChanged(NameOf(CanGenerateImages))
                    OnPropertyChanged(NameOf(ImageButtonText))
                    OnPropertyChanged(NameOf(ImageButtonBackground))
                    OnPropertyChanged(NameOf(ProjectStageText))
                    OnPropertyChanged(NameOf(LocalizedStageText))
                    OnPropertyChanged(NameOf(StageBrush))
                    OnPropertyChanged(NameOf(ProgressPercent))
                    OnPropertyChanged(NameOf(ProgressLabelText))
                    NotifyProjectStateBadgesChanged()
                    NotifyStepBrushesChanged()
                End If
            End Set
        End Property

        Public Property HasGeneratedPlan As Boolean
            Get
                Return _hasGeneratedPlan
            End Get
            Set(value As Boolean)
                If SetProperty(_hasGeneratedPlan, value) Then
                    OnPropertyChanged(NameOf(CanGenerateImages))
                    OnPropertyChanged(NameOf(CanRegenerateSelectedImage))
                    OnPropertyChanged(NameOf(PlanButtonText))
                    OnPropertyChanged(NameOf(PlanButtonBackground))
                    OnPropertyChanged(NameOf(ProjectStageText))
                    OnPropertyChanged(NameOf(LocalizedStageText))
                    OnPropertyChanged(NameOf(StageBrush))
                    OnPropertyChanged(NameOf(ProgressPercent))
                    OnPropertyChanged(NameOf(ProgressLabelText))
                    NotifyProjectStateBadgesChanged()
                    NotifyStepBrushesChanged()
                End If
            End Set
        End Property

        Public Property LastOutputDirectory As String
            Get
                Return _lastOutputDirectory
            End Get
            Set(value As String)
                If SetProperty(_lastOutputDirectory, value) Then
                    OnPropertyChanged(NameOf(CanOpenOutputDirectory))
                    ConfigureOutputWatcher(value)
                End If
            End Set
        End Property

        Public Property SelectedGeneratedImage As GeneratedImageItem
            Get
                Return _selectedGeneratedImage
            End Get
            Set(value As GeneratedImageItem)
                If SetProperty(_selectedGeneratedImage, value) Then
                    OnPropertyChanged(NameOf(HasSelectedGeneratedImage))
                    OnPropertyChanged(NameOf(CanPreviewSelectedImage))
                    OnPropertyChanged(NameOf(CanRegenerateSelectedImage))
                    OnPropertyChanged(NameOf(CanEditSelectedImage))
                    OnPropertyChanged(NameOf(CanDownloadSelectedImage))
                End If
            End Set
        End Property

        Public ReadOnly Property BusyVisibility As Visibility
            Get
                Return If(IsBusy, Visibility.Visible, Visibility.Collapsed)
            End Get
        End Property

        Public ReadOnly Property CanGeneratePlan As Boolean
            Get
                If IsPlanning Then Return True
                Return UploadedImages.Count > 0 AndAlso Not IsBusy
            End Get
        End Property

        Public ReadOnly Property CanGenerateImages As Boolean
            Get
                If IsGeneratingImages Then Return True
                Return UploadedImages.Count > 0 AndAlso HasGeneratedPlan AndAlso Not IsBusy
            End Get
        End Property

        Public ReadOnly Property PlanButtonText As String
            Get
                If IsPlanning Then Return "Stop Planning"
                If HasGeneratedPlan Then Return "Regenerate Plan"
                Return "Generate Plan"
            End Get
        End Property

        Public ReadOnly Property ImageButtonText As String
            Get
                If IsGeneratingImages Then Return "Stop Rendering"
                Return "Generate Images"
            End Get
        End Property

        Public Sub NotifyLocalizationChanged()
            OnPropertyChanged(NameOf(LocalizedStageText))
            OnPropertyChanged(NameOf(FailureSummaryText))
            OnPropertyChanged(NameOf(LastUpdatedText))
            OnPropertyChanged(NameOf(QueueRemainingText))
            OnPropertyChanged(NameOf(QueueActionText))
            OnPropertyChanged(NameOf(ResolutionHintText))
            NotifyProjectStateBadgesChanged()
        End Sub

        Public ReadOnly Property PlanButtonBackground As Brush
            Get
                If IsPlanning Then Return StopActionBrush
                If HasGeneratedPlan Then Return CompletedActionBrush
                Return If(CanGeneratePlan, PrimaryActionBrush, DisabledActionBrush)
            End Get
        End Property

        Public ReadOnly Property ImageButtonBackground As Brush
            Get
                If IsGeneratingImages Then Return StopActionBrush
                Return If(CanGenerateImages, PrimaryActionBrush, DisabledActionBrush)
            End Get
        End Property

        Public ReadOnly Property CanOpenOutputDirectory As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(LastOutputDirectory) AndAlso IO.Directory.Exists(LastOutputDirectory)
            End Get
        End Property

        Public ReadOnly Property HasSelectedGeneratedImage As Boolean
            Get
                Return SelectedGeneratedImage IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property CanPreviewSelectedImage As Boolean
            Get
                Return SelectedGeneratedImage IsNot Nothing AndAlso SelectedGeneratedImage.HasPreview AndAlso Not IsBusy
            End Get
        End Property

        Public ReadOnly Property CanEditSelectedImage As Boolean
            Get
                Return SelectedGeneratedImage IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(SelectedGeneratedImage.SavedPath) AndAlso File.Exists(SelectedGeneratedImage.SavedPath)
            End Get
        End Property

        Public ReadOnly Property CanDownloadSelectedImage As Boolean
            Get
                Return CanEditSelectedImage
            End Get
        End Property

        Public ReadOnly Property CanRegenerateSelectedImage As Boolean
            Get
                Return SelectedGeneratedImage IsNot Nothing AndAlso UploadedImages.Count > 0 AndAlso HasGeneratedPlan AndAlso Not IsBusy
            End Get
        End Property

        Public ReadOnly Property LoadedImageCountText As String
            Get
                Return $"{UploadedImages.Count}/10"
            End Get
        End Property
        Public Sub SetUploadedImages(filePaths As IEnumerable(Of String))
            If filePaths Is Nothing Then Return

            Dim existingPaths = New HashSet(Of String)(
                UploadedImages.Select(Function(item) item.FullPath),
                StringComparer.OrdinalIgnoreCase)

            Dim addedCount = 0
            For Each filePath In filePaths.
                Where(Function(path) Not String.IsNullOrWhiteSpace(path)).
                Select(Function(path) path.Trim()).
                Where(Function(path) Not existingPaths.Contains(path)).
                Take(Math.Max(0, 10 - UploadedImages.Count))

                UploadedImages.Add(New ProductImageItem With {
                    .FullPath = filePath,
                    .FileName = Path.GetFileName(filePath),
                    .Preview = LoadBitmap(filePath)
                })
                existingPaths.Add(filePath)
                addedCount += 1
            Next

            WriteLog($"Import images: project={ProjectName}, count={UploadedImages.Count}, added={addedCount}, files={String.Join(", ", UploadedImages.Select(Function(item) item.FileName))}")
            RefreshUploadedImageState($"Imported {addedCount} image(s). Total: {UploadedImages.Count}.")
        End Sub

        Public Sub MarkActivated()
            LastActivatedAt = DateTime.Now
        End Sub

        Public Function ToPersistedRecord() As PersistedProjectRecord
            Return New PersistedProjectRecord With {
                .ProjectId = ProjectId,
                .ProjectName = ProjectName,
                .IsArchived = IsArchived,
                .ProjectSummary = ProjectSummary,
                .SelectedScenarioPresetId = SelectedScenarioPresetId,
                .SelectedPlannerProviderKey = SelectedPlannerProviderKey,
                .RequirementText = RequirementText,
                .TargetLanguage = TargetLanguage,
                .SelectedPlannerModel = SelectedPlannerModel,
                .SelectedImageProviderKey = SelectedImageProviderKey,
                .SelectedImageModel = SelectedImageModel,
                .SelectedAspectRatio = SelectedAspectRatio,
                .OutputResolution = OutputResolution,
                .RequestedCount = RequestedCount,
                .DerivativeImageCount = DerivativeImageCount,
                .AutoRenderAfterPlan = AutoRenderAfterPlan,
                .DesignPlanMarkdown = DesignPlanMarkdown,
                .ImagePlanMarkdown = ImagePlanMarkdown,
                .HasGeneratedPlan = HasGeneratedPlan,
                .LastOutputDirectory = LastOutputDirectory,
                .LastActivatedAt = LastActivatedAt,
                .UploadedImagePaths = UploadedImages.Select(Function(item) item.FullPath).Where(Function(path) Not String.IsNullOrWhiteSpace(path)).ToList(),
                .GeneratedImages = GeneratedImages.Select(Function(item) New PersistedGeneratedImageRecord With {
                    .Sequence = item.Sequence,
                    .Title = item.Title,
                    .Description = item.Description,
                    .SavedPath = item.SavedPath,
                    .StatusText = item.StatusText,
                    .LastError = item.LastError
                }).ToList()
            }
        End Function

        Public Sub RestoreFromRecord(record As PersistedProjectRecord)
            If record Is Nothing Then Return

            ProjectName = If(String.IsNullOrWhiteSpace(record.ProjectName), ProjectName, record.ProjectName)
            IsArchived = record.IsArchived
            ProjectSummary = If(String.IsNullOrWhiteSpace(record.ProjectSummary), ProjectSummary, record.ProjectSummary)
            SelectedScenarioPresetId = record.SelectedScenarioPresetId
            SelectedPlannerProviderKey = record.SelectedPlannerProviderKey
            RequirementText = record.RequirementText
            TargetLanguage = If(String.IsNullOrWhiteSpace(record.TargetLanguage), NoTextLanguage, record.TargetLanguage)
            SelectedPlannerModel = If(String.IsNullOrWhiteSpace(record.SelectedPlannerModel), SelectedPlannerModel, record.SelectedPlannerModel)
            SelectedImageProviderKey = record.SelectedImageProviderKey
            SelectedImageModel = If(String.IsNullOrWhiteSpace(record.SelectedImageModel), SelectedImageModel, record.SelectedImageModel)
            SelectedAspectRatio = If(String.IsNullOrWhiteSpace(record.SelectedAspectRatio), "1:1", record.SelectedAspectRatio)
            OutputResolution = If(String.IsNullOrWhiteSpace(record.OutputResolution), OutputResolution, record.OutputResolution)
            RequestedCount = If(record.RequestedCount <= 0, 4, record.RequestedCount)
            DerivativeImageCount = If(record.DerivativeImageCount <= 0, 1, record.DerivativeImageCount)
            AutoRenderAfterPlan = record.AutoRenderAfterPlan
            DesignPlanMarkdown = record.DesignPlanMarkdown
            ImagePlanMarkdown = record.ImagePlanMarkdown
            HasGeneratedPlan = record.HasGeneratedPlan
            LastActivatedAt = If(record.LastActivatedAt = DateTime.MinValue, DateTime.Now, record.LastActivatedAt)
            RefreshProviderConfiguration()
            EnsureProjectSummaryFromRequirement()

            UploadedImages.Clear()
            For Each filePath In record.UploadedImagePaths.Where(Function(path) Not String.IsNullOrWhiteSpace(path))
                UploadedImages.Add(New ProductImageItem With {
                    .FullPath = filePath,
                    .FileName = Path.GetFileName(filePath),
                    .Preview = If(File.Exists(filePath), LoadBitmap(filePath), Nothing)
                })
            Next

            GeneratedImages.Clear()
            For Each item In record.GeneratedImages.OrderBy(Function(entry) entry.Sequence)
                GeneratedImages.Add(New GeneratedImageItem With {
                    .Sequence = item.Sequence,
                    .Title = item.Title,
                    .Description = item.Description,
                    .SavedPath = item.SavedPath,
                    .Preview = If(Not String.IsNullOrWhiteSpace(item.SavedPath) AndAlso File.Exists(item.SavedPath), LoadBitmap(item.SavedPath), Nothing),
                    .StatusText = item.StatusText,
                    .StatusBrush = ResolveStatusBrush(item.StatusText),
                    .LastError = item.LastError
                })
            Next

            Dim normalizedOutputDirectory = record.LastOutputDirectory
            If String.IsNullOrWhiteSpace(normalizedOutputDirectory) AndAlso GeneratedImages.Count > 0 Then
                normalizedOutputDirectory = Path.GetDirectoryName(GeneratedImages.First().SavedPath)
            End If
            LastOutputDirectory = If(normalizedOutputDirectory, String.Empty)

            RefreshRestoredImageBadges()
            SelectedGeneratedImage = GeneratedImages.OrderBy(Function(item) item.Sequence).FirstOrDefault()
            StatusMessage = F("Project.SelectedModels", "Planner: {0} | Image: {1}", SelectedPlannerDisplayName, SelectedImageDisplayName)
            ExecutionSummary = If(GeneratedImages.Count > 0,
                                  $"已恢复 {GeneratedImages.Count} 张输出记录。",
                                  DefaultExecutionSummary)
            OnPropertyChanged(NameOf(CanGeneratePlan))
            OnPropertyChanged(NameOf(CanGenerateImages))
            OnPropertyChanged(NameOf(ProjectCover))
            OnPropertyChanged(NameOf(HasProjectCover))
            OnPropertyChanged(NameOf(LoadedImageCountText))
            OnPropertyChanged(NameOf(SourceSummaryText))
            OnPropertyChanged(NameOf(ProjectSubtitle))
            NotifyProjectStateBadgesChanged()
            NotifyStepBrushesChanged()
        End Sub

        Public Sub RemoveUploadedImage(item As ProductImageItem)
            If item Is Nothing Then Return

            UploadedImages.Remove(item)
            WriteLog($"Remove image: project={ProjectName}, file={item.FileName}")
            RefreshUploadedImageState(If(UploadedImages.Count = 0,
                                         "All source images cleared.",
                                         $"Removed one source image. Remaining: {UploadedImages.Count}."))
        End Sub

        Public Async Function ToggleGeneratePlanAsync() As Task
            If IsPlanning Then
                StopCurrentOperation("Stopping planning...")
                Return
            End If

            If UploadedImages.Count = 0 Then
                Throw New InvalidOperationException("Please import at least one product image first.")
            End If

            Dim cancellationToken = BeginOperation(True, $"Planning with {SelectedPlannerDisplayName}...")
            WriteLog($"Start plan: project={ProjectName}, scene={SelectedScenarioPresetName}, provider={SelectedPlannerDisplayName}, imageCount={UploadedImages.Count}, requestedCount={RequestedCount}, language={TargetLanguage}, aspectRatio={If(IsCustomImageSizeMode, OutputResolution.Replace("*", ":"), SelectedAspectRatio)}")
            Dim shouldAutoRender As Boolean = False

            Try
                Dim usesEnglishPlanning = String.Equals(LocalizationManager.Instance.CurrentLanguageCode, "en", StringComparison.OrdinalIgnoreCase)
                Dim request As New AnalyzeRequest With {
                    .RequirementText = RequirementText,
                    .TargetLanguage = TargetLanguage,
                    .PlanningLanguageCode = LocalizationManager.Instance.CurrentLanguageCode,
                    .AspectRatio = If(IsCustomImageSizeMode, NormalizeCustomAspectRatio(OutputResolution), NormalizeAspectRatio(SelectedAspectRatio)),
                    .RequestedCount = RequestedCount,
                    .ScenarioName = SelectedScenarioPresetName,
                    .ScenarioDescription = SelectedScenarioPresetDescription,
                    .ScenarioDesignPlanningTemplate = If(GetSelectedScenarioPreset()?.DesignPlanningTemplate, String.Empty),
                    .ScenarioImagePlanningTemplate = If(GetSelectedScenarioPreset()?.ImagePlanningTemplate, String.Empty)
                }

                Dim designPlan = Await _vertexService.AnalyzeProductAsync(UploadedImages, request, SelectedPlannerProviderKey, SelectedPlannerModel, cancellationToken, AddressOf WriteLog)
                UpdateProjectSummaryFromPlan(designPlan)
                DesignPlanMarkdown = BuildDesignPlanMarkdown(designPlan, usesEnglishPlanning)
                ImagePlanMarkdown = BuildImagePlanMarkdown(designPlan.ImagePlans, TargetLanguage, usesEnglishPlanning)
                HasGeneratedPlan = True
                StatusMessage = "Plan generated. You can edit markdown before rendering."
                WriteLog($"Plan complete: project={ProjectName}, imagePlanCount={designPlan.ImagePlans.Count}")
                shouldAutoRender = AutoRenderAfterPlan
            Catch ex As OperationCanceledException
                StatusMessage = "Planning stopped."
                WriteLog($"Plan stopped: project={ProjectName}")
            Finally
                EndOperation()
            End Try

            If shouldAutoRender Then
                StatusMessage = "Plan generated. Auto rendering started."
                WriteLog($"Auto render triggered: project={ProjectName}")
                Await ToggleGenerateImagesAsync()
            End If
        End Function

        Public Async Function ToggleGenerateImagesAsync() As Task
            If IsGeneratingImages Then
                StopCurrentOperation("Stopping image generation...")
                Return
            End If

            If UploadedImages.Count = 0 Then
                Throw New InvalidOperationException("Please import product images first.")
            End If

            Dim imagePlans = GetEffectiveImagePlans()
            If imagePlans.Count = 0 Then
                Throw New InvalidOperationException("Image plan markdown could not be parsed.")
            End If

            Dim outputDirectory = PrepareOutputDirectory(True)
            Dim cancellationToken = BeginOperation(False, "Generating images...")
            WriteLog($"Start render batch: project={ProjectName}, scene={SelectedScenarioPresetName}, count={imagePlans.Count}, outputDirectory={outputDirectory}, resolution={OutputResolution}, aspectRatio={SelectedAspectRatio}, language={TargetLanguage}")

            Try
                LastOutputDirectory = outputDirectory
                GeneratedImages.Clear()

                For Each plan In imagePlans.OrderBy(Function(item) item.Sequence)
                    GeneratedImages.Add(BuildPendingImageItem(plan, outputDirectory))
                Next

                SelectedGeneratedImage = GeneratedImages.FirstOrDefault()

                For Each plan In imagePlans.OrderBy(Function(item) item.Sequence)
                    cancellationToken.ThrowIfCancellationRequested()
                    Dim currentItem = GeneratedImages.First(Function(item) item.Sequence = plan.Sequence)
                    SelectedGeneratedImage = currentItem
                    Await GenerateSinglePlanAsync(plan, outputDirectory, currentItem, cancellationToken)
                Next

                Dim successCount = GeneratedImages.Where(Function(item) String.Equals(item.StatusText, "Success", StringComparison.Ordinal)).Count()
                Dim failedCount = GeneratedImages.Where(Function(item) String.Equals(item.StatusText, "Failed", StringComparison.Ordinal)).Count()
                ExecutionSummary = $"Rendered {imagePlans.Count} item(s). Success: {successCount}. Failed: {failedCount}."
                StatusMessage = If(failedCount = 0, "Image generation complete.", $"Image generation complete with {failedCount} failure(s).")
                WriteLog($"Render batch complete: project={ProjectName}, success={successCount}, failed={failedCount}, outputDirectory={outputDirectory}")
                RefreshGalleryFromOutputDirectory()
            Catch ex As OperationCanceledException
                For Each pendingItem In GeneratedImages.Where(Function(item) item.StatusText = "Pending" OrElse item.StatusText = "Running")
                    pendingItem.StatusText = "Stopped"
                    pendingItem.StatusBrush = PendingBrush
                Next

                StatusMessage = "Image generation stopped."
                ExecutionSummary = $"Stopped. Generated {GeneratedImages.Where(Function(item) item.StatusText = "Success").Count()} item(s)."
                WriteLog($"Render batch stopped: project={ProjectName}")
            Finally
                EndOperation()
            End Try
        End Function

        Public Async Function RegenerateImageAsync(item As GeneratedImageItem) As Task
            If item Is Nothing Then Return
            If UploadedImages.Count = 0 Then Throw New InvalidOperationException("Please import product images first.")

            Dim imagePlans = GetEffectiveImagePlans()
            Dim plan = imagePlans.FirstOrDefault(Function(entry) entry.Sequence = item.Sequence)
            If plan Is Nothing Then
                plan = imagePlans.FirstOrDefault(Function(entry) String.Equals(entry.Title, item.Title, StringComparison.OrdinalIgnoreCase))
            End If
            If plan Is Nothing Then
                Throw New InvalidOperationException("Could not find the matching plan for this image in the current markdown.")
            End If

            Dim outputDirectory = PrepareOutputDirectory(False)
            Dim cancellationToken = BeginOperation(False, $"Regenerating image {plan.Sequence}...")
            WriteLog($"Start regenerate: project={ProjectName}, sequence={plan.Sequence}, title={plan.Title}, outputDirectory={outputDirectory}")

            Try
                LastOutputDirectory = outputDirectory
                SelectedGeneratedImage = item
                Await GenerateSinglePlanAsync(plan, outputDirectory, item, cancellationToken)
                ExecutionSummary = $"Regenerated image {plan.Sequence}. Gallery count: {GeneratedImages.Count}."
                StatusMessage = $"Regeneration complete: image {plan.Sequence}."
                WriteLog($"Regenerate complete: project={ProjectName}, sequence={plan.Sequence}, title={plan.Title}")
                RefreshGalleryFromOutputDirectory()
            Catch ex As OperationCanceledException
                item.StatusText = "Stopped"
                item.StatusBrush = PendingBrush
                StatusMessage = $"Regeneration stopped: image {plan.Sequence}."
                WriteLog($"Regenerate stopped: project={ProjectName}, sequence={plan.Sequence}, title={plan.Title}")
            Finally
                EndOperation()
            End Try
        End Function

        Public Sub StopCurrentOperation(statusText As String)
            If _operationCancellationTokenSource Is Nothing OrElse _operationCancellationTokenSource.IsCancellationRequested Then
                Return
            End If

            StatusMessage = statusText
            WriteLog($"Stop requested: project={ProjectName}, status={statusText}")
            _operationCancellationTokenSource.Cancel()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            StopCurrentOperation("Closing project...")
            ConfigureOutputWatcher(String.Empty)
            _operationCancellationTokenSource?.Dispose()
            _refreshTimer.Stop()
        End Sub

        Private Function BeginOperation(isPlan As Boolean, statusText As String) As CancellationToken
            _operationCancellationTokenSource?.Dispose()
            _operationCancellationTokenSource = New CancellationTokenSource()

            IsPlanning = isPlan
            IsGeneratingImages = Not isPlan
            IsBusy = True
            StatusMessage = statusText
            Return _operationCancellationTokenSource.Token
        End Function

        Private Sub EndOperation()
            IsPlanning = False
            IsGeneratingImages = False
            IsBusy = False
            _operationCancellationTokenSource?.Dispose()
            _operationCancellationTokenSource = Nothing
        End Sub

        Private Async Function GenerateSinglePlanAsync(plan As ImagePlanItem, outputDirectory As String, targetItem As GeneratedImageItem, cancellationToken As CancellationToken) As Task
            targetItem.StatusText = "Running"
            targetItem.StatusBrush = RunningBrush
            targetItem.LastError = String.Empty
            WriteLog($"Render image start: project={ProjectName}, sequence={plan.Sequence}, title={plan.Title}, outputFile={BuildOutputFilePath(outputDirectory, plan)}")

            For attempt = 1 To MaxImageGenerationAttempts
                cancellationToken.ThrowIfCancellationRequested()
                Dim retryDelay As Integer = 0
                Dim shouldStopRetry As Boolean = False

                Try
                    WriteLog($"Render image attempt: project={ProjectName}, sequence={plan.Sequence}, title={plan.Title}, attempt={attempt}/{MaxImageGenerationAttempts}")

                    Dim results = Await _vertexService.GenerateImagesAsync(
                        UploadedImages,
                        {plan},
                        SelectedImageProviderKey,
                        SelectedImageModel,
                        NormalizeAspectRatio(SelectedAspectRatio),
                        NormalizeImageSize(OutputResolution),
                        TargetLanguage,
                        SelectedScenarioPresetName,
                        If(GetSelectedScenarioPreset()?.ImageInstruction, String.Empty),
                        outputDirectory,
                        Nothing,
                        cancellationToken,
                        AddressOf WriteLog)

                    Dim generated = results.FirstOrDefault()
                    If generated Is Nothing OrElse String.IsNullOrWhiteSpace(generated.SavedPath) Then
                        Throw New InvalidOperationException($"Model returned no image bytes for {plan.Title}.")
                    End If

                    targetItem.Sequence = plan.Sequence
                    targetItem.Title = generated.Title
                    targetItem.Description = generated.Description
                    targetItem.SavedPath = generated.SavedPath
                    targetItem.Preview = generated.Preview
                    targetItem.StatusText = "Success"
                    targetItem.StatusBrush = SuccessBrush
                    targetItem.LastError = String.Empty
                    WriteLog($"Render image success: project={ProjectName}, sequence={plan.Sequence}, savedPath={generated.SavedPath}, attempt={attempt}/{MaxImageGenerationAttempts}")
                    Return
                Catch ex As OperationCanceledException
                    targetItem.StatusText = "Stopped"
                    targetItem.StatusBrush = PendingBrush
                    WriteLog($"Render image stopped: project={ProjectName}, sequence={plan.Sequence}, title={plan.Title}, attempt={attempt}/{MaxImageGenerationAttempts}")
                    Throw
                Catch ex As Exception
                    targetItem.Sequence = plan.Sequence
                    targetItem.Title = plan.Title
                    targetItem.Description = plan.Purpose
                    targetItem.SavedPath = BuildOutputFilePath(outputDirectory, plan)
                    targetItem.LastError = ex.Message

                    If attempt < MaxImageGenerationAttempts Then
                        retryDelay = Random.Shared.Next(ImageRetryDelayMinMilliseconds, ImageRetryDelayMaxMilliseconds + 1)
                        targetItem.StatusText = $"Retry {attempt + 1}/{MaxImageGenerationAttempts}"
                        targetItem.StatusBrush = MissingBrush
                        WriteLog($"Render image retry scheduled: project={ProjectName}, sequence={plan.Sequence}, title={plan.Title}, nextAttempt={attempt + 1}/{MaxImageGenerationAttempts}, delayMs={retryDelay}", ex)
                    Else
                        targetItem.StatusText = "Failed"
                        targetItem.StatusBrush = FailedBrush
                        WriteLog($"Render image failed: project={ProjectName}, sequence={plan.Sequence}, title={plan.Title}, attempts={MaxImageGenerationAttempts}", ex)
                        shouldStopRetry = True
                    End If
                End Try

                If retryDelay > 0 Then
                    Await Task.Delay(retryDelay, cancellationToken)
                    Continue For
                End If

                If shouldStopRetry Then
                    Exit For
                End If
            Next
        End Function
        Private Function GetEffectiveImagePlans() As List(Of ImagePlanItem)
            Dim basePlans = ParseImagePlanMarkdown(ImagePlanMarkdown).
                OrderBy(Function(item) item.Sequence).
                Take(RequestedCount).
                ToList()

            Return ExpandImagePlans(basePlans, DerivativeImageCount)
        End Function

        Private Sub RefreshTimer_Tick(sender As Object, e As EventArgs)
            _refreshTimer.Stop()
            RefreshGalleryFromOutputDirectory()
        End Sub

        Private Sub GeneratedImages_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs)
            If SelectedGeneratedImage Is Nothing AndAlso GeneratedImages.Count > 0 Then
                SelectedGeneratedImage = GeneratedImages.FirstOrDefault()
            ElseIf SelectedGeneratedImage IsNot Nothing AndAlso Not GeneratedImages.Contains(SelectedGeneratedImage) Then
                SelectedGeneratedImage = GeneratedImages.FirstOrDefault()
            End If

            OnPropertyChanged(NameOf(StageBrush))
            OnPropertyChanged(NameOf(ProgressPercent))
            OnPropertyChanged(NameOf(ProgressLabelText))
            OnPropertyChanged(NameOf(ProjectCover))
            OnPropertyChanged(NameOf(HasProjectCover))
            OnPropertyChanged(NameOf(AssetCountText))
            NotifyProjectStateBadgesChanged()
            NotifyStepBrushesChanged()
        End Sub

        Private Sub OutputWatcher_Changed(sender As Object, e As FileSystemEventArgs)
            If Not IsSupportedImageFile(e.FullPath) Then Return

            _dispatcher.Invoke(Sub()
                                   _refreshTimer.Stop()
                                   _refreshTimer.Start()
                               End Sub)
        End Sub

        Private Sub ConfigureOutputWatcher(directoryPath As String)
            If _outputWatcher IsNot Nothing Then
                RemoveHandler _outputWatcher.Created, AddressOf OutputWatcher_Changed
                RemoveHandler _outputWatcher.Changed, AddressOf OutputWatcher_Changed
                RemoveHandler _outputWatcher.Deleted, AddressOf OutputWatcher_Changed
                RemoveHandler _outputWatcher.Renamed, AddressOf OutputWatcher_Changed
                _outputWatcher.Dispose()
                _outputWatcher = Nothing
            End If

            If String.IsNullOrWhiteSpace(directoryPath) OrElse Not IO.Directory.Exists(directoryPath) Then Return

            _outputWatcher = New FileSystemWatcher(directoryPath) With {
                .Filter = "*.*",
                .IncludeSubdirectories = False,
                .NotifyFilter = NotifyFilters.FileName Or NotifyFilters.LastWrite Or NotifyFilters.CreationTime Or NotifyFilters.Size,
                .EnableRaisingEvents = True
            }

            AddHandler _outputWatcher.Created, AddressOf OutputWatcher_Changed
            AddHandler _outputWatcher.Changed, AddressOf OutputWatcher_Changed
            AddHandler _outputWatcher.Deleted, AddressOf OutputWatcher_Changed
            AddHandler _outputWatcher.Renamed, AddressOf OutputWatcher_Changed
        End Sub

        Private Sub RefreshGalleryFromOutputDirectory()
            If Not CanOpenOutputDirectory Then Return

            Dim imageFiles = IO.Directory.EnumerateFiles(LastOutputDirectory).
                Where(AddressOf IsSupportedImageFile).
                OrderBy(AddressOf GetSortableFileName).
                ToList()

            For Each item In GeneratedImages.ToList()
                If String.IsNullOrWhiteSpace(item.SavedPath) Then Continue For

                If imageFiles.Any(Function(path) String.Equals(path, item.SavedPath, StringComparison.OrdinalIgnoreCase)) Then
                    item.Preview = LoadBitmap(item.SavedPath)
                    item.StatusText = "Success"
                    item.StatusBrush = SuccessBrush
                    item.LastError = String.Empty
                ElseIf String.Equals(item.StatusText, "Success", StringComparison.Ordinal) Then
                    item.Preview = Nothing
                    item.StatusText = "Missing"
                    item.StatusBrush = MissingBrush
                End If
            Next

            Dim overflowItems = GeneratedImages.OrderBy(Function(item) item.Sequence).Skip(EffectiveRenderItemCount).ToList()
            For Each overflowItem In overflowItems
                GeneratedImages.Remove(overflowItem)
            Next

            If SelectedGeneratedImage Is Nothing OrElse Not GeneratedImages.Contains(SelectedGeneratedImage) Then
                SelectedGeneratedImage = GeneratedImages.OrderBy(Function(item) item.Sequence).FirstOrDefault()
            End If
        End Sub

        Private Sub RefreshUploadedImageState(statusText As String)
            Dim index = 1
            For Each item In UploadedImages
                item.BadgeText = F("Project.Badge.Image", "Img {0}", index)
                index += 1
            Next

            HasGeneratedPlan = False
            DesignPlanMarkdown = String.Empty
            ImagePlanMarkdown = String.Empty
            LastOutputDirectory = String.Empty
            GeneratedImages.Clear()
            SelectedGeneratedImage = Nothing
            StatusMessage = statusText
            ExecutionSummary = DefaultExecutionSummary

            OnPropertyChanged(NameOf(LoadedImageCountText))
            OnPropertyChanged(NameOf(CanGeneratePlan))
            OnPropertyChanged(NameOf(CanGenerateImages))
            OnPropertyChanged(NameOf(CanRegenerateSelectedImage))
            OnPropertyChanged(NameOf(PlanButtonBackground))
            OnPropertyChanged(NameOf(PlanButtonText))
            OnPropertyChanged(NameOf(ImageButtonBackground))
            OnPropertyChanged(NameOf(ImageButtonText))
            OnPropertyChanged(NameOf(ProjectStageText))
            OnPropertyChanged(NameOf(LocalizedStageText))
            OnPropertyChanged(NameOf(ProjectCover))
            OnPropertyChanged(NameOf(HasProjectCover))
            OnPropertyChanged(NameOf(ProjectSubtitle))
            OnPropertyChanged(NameOf(AssetCountText))
            OnPropertyChanged(NameOf(SourceSummaryText))
            OnPropertyChanged(NameOf(ProgressPercent))
            OnPropertyChanged(NameOf(ProgressLabelText))
            NotifyProjectStateBadgesChanged()
            NotifyStepBrushesChanged()
        End Sub

        Private Sub RefreshRestoredImageBadges()
            Dim index = 1
            For Each item In UploadedImages
                item.BadgeText = F("Project.Badge.Image", "Img {0}", index)
                index += 1
            Next
        End Sub

        Private Sub EnsureProjectSummaryFromRequirement(Optional force As Boolean = False)
            If Not force AndAlso Not IsDefaultProjectSummary(_projectSummary) Then
                Return
            End If

            Dim fallbackSummary = BuildFallbackProjectSummary(_requirementText)
            If String.IsNullOrWhiteSpace(fallbackSummary) Then
                If force Then
                    ProjectSummary = DefaultProjectSummaryText
                End If
                Return
            End If

            ProjectSummary = fallbackSummary
        End Sub

        Private Sub UpdateProjectSummaryFromPlan(plan As DesignPlan)
            If plan Is Nothing Then
                EnsureProjectSummaryFromRequirement()
                Return
            End If

            Dim generatedSummary = NormalizeSummaryText(plan.ProductSummary)
            If String.IsNullOrWhiteSpace(generatedSummary) Then
                EnsureProjectSummaryFromRequirement()
                Return
            End If

            Dim fallbackSummary = BuildFallbackProjectSummary(_requirementText)
            If IsDefaultProjectSummary(_projectSummary) OrElse
               String.Equals(_projectSummary, fallbackSummary, StringComparison.Ordinal) Then
                ProjectSummary = generatedSummary
            End If
        End Sub

        Private Sub NotifyProjectStateBadgesChanged()
            OnPropertyChanged(NameOf(QueueRemainingText))
            OnPropertyChanged(NameOf(QueueActionText))
            OnPropertyChanged(NameOf(QueueActionBackground))
            OnPropertyChanged(NameOf(QueueActionBorderBrush))
            OnPropertyChanged(NameOf(QueueActionForeground))
            OnPropertyChanged(NameOf(ProgressBrush))
            OnPropertyChanged(NameOf(SuccessfulImageCount))
            OnPropertyChanged(NameOf(RetryableFailureCount))
            OnPropertyChanged(NameOf(HasRetryableFailures))
            OnPropertyChanged(NameOf(FailureSummaryText))
            OnPropertyChanged(NameOf(FailureVisibility))
            OnPropertyChanged(NameOf(LastUpdatedText))
            OnPropertyChanged(NameOf(ScenarioSummaryText))
        End Sub

        Private Function GetSelectedScenarioPreset() As ScenarioPreset
            If _availableScenarioPresets Is Nothing OrElse _availableScenarioPresets.Count = 0 Then
                Return Nothing
            End If

            Dim selected = _availableScenarioPresets.FirstOrDefault(Function(item) String.Equals(item.Id, _selectedScenarioPresetId, StringComparison.Ordinal))
            If selected IsNot Nothing Then
                Return selected
            End If

            Return _availableScenarioPresets.FirstOrDefault()
        End Function

        Private Sub NotifyStepBrushesChanged()
            OnPropertyChanged(NameOf(StepInputBrush))
            OnPropertyChanged(NameOf(StepPlanBrush))
            OnPropertyChanged(NameOf(StepRenderBrush))
            OnPropertyChanged(NameOf(StepReviewBrush))
        End Sub

        Private Function PrepareOutputDirectory(createNewBatch As Boolean) As String
            Dim outputsRoot = Path.Combine(AppPaths.OutputsRoot, SanitizeFileName(ProjectName))
            IO.Directory.CreateDirectory(outputsRoot)

            If createNewBatch OrElse String.IsNullOrWhiteSpace(LastOutputDirectory) OrElse Not IO.Directory.Exists(LastOutputDirectory) Then
                Dim folderName = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff", CultureInfo.InvariantCulture)
                Dim candidate = Path.Combine(outputsRoot, folderName)
                Dim suffix = 1

                While IO.Directory.Exists(candidate)
                    candidate = Path.Combine(outputsRoot, $"{folderName}_{suffix}")
                    suffix += 1
                End While

                IO.Directory.CreateDirectory(candidate)
                Return candidate
            End If

            Return LastOutputDirectory
        End Function

        Private Shared Function BuildPendingImageItem(plan As ImagePlanItem, outputDirectory As String) As GeneratedImageItem
            Return New GeneratedImageItem With {
                .Sequence = plan.Sequence,
                .Title = plan.Title,
                .Description = plan.Purpose,
                .SavedPath = BuildOutputFilePath(outputDirectory, plan),
                .StatusText = "Pending",
                .StatusBrush = PendingBrush
            }
        End Function

        Private Shared Function ResolveStatusBrush(statusText As String) As Brush
            Select Case statusText
                Case "Success"
                    Return SuccessBrush
                Case "Failed"
                    Return FailedBrush
                Case "Missing"
                    Return MissingBrush
                Case "Running"
                    Return RunningBrush
                Case Else
                    Return PendingBrush
            End Select
        End Function

        Private Shared Function BuildOutputFilePath(outputDirectory As String, plan As ImagePlanItem) As String
            Return Path.Combine(outputDirectory, $"{plan.Sequence:00}_{SanitizeFileName(plan.Title)}.png")
        End Function

        Private Overloads Function WriteLog(message As String) As String
            Return AppendRecentLog(_logger.Info(message))
        End Function

        Private Overloads Function WriteLog(message As String, ex As Exception) As String
            Return AppendRecentLog(_logger.Error(message, ex))
        End Function

        Private Function AppendRecentLog(entry As String) As String
            If _dispatcher.CheckAccess() Then
                UpdateRecentLogText(entry)
            Else
                _dispatcher.Invoke(Sub() UpdateRecentLogText(entry))
            End If

            Return entry
        End Function

        Private Sub UpdateRecentLogText(entry As String)
            Dim existingLines = RecentLogText.Split({Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries).ToList()
            existingLines.Add(entry)
            If existingLines.Count > 120 Then
                existingLines = existingLines.Skip(existingLines.Count - 120).ToList()
            End If
            RecentLogText = String.Join(Environment.NewLine, existingLines)
            LastActivatedAt = DateTime.Now
        End Sub

        Private Shared Function BuildDesignPlanMarkdown(plan As DesignPlan, useEnglish As Boolean) As String
            Dim builder As New StringBuilder()

            If useEnglish Then
                builder.AppendLine("# Design Plan")
                builder.AppendLine()
                builder.AppendLine("## Product Summary")
                builder.AppendLine(plan.ProductSummary)
                builder.AppendLine()
                builder.AppendLine("## Design Theme")
                builder.AppendLine(plan.DesignTheme)
                builder.AppendLine()
                builder.AppendLine("## Audience")
                builder.AppendLine(plan.Audience)
                builder.AppendLine()
                builder.AppendLine("## Color System")
                builder.AppendLine(plan.ColorSystem)
                builder.AppendLine()
                builder.AppendLine("## Typography")
                builder.AppendLine(plan.Typography)
                builder.AppendLine()
                builder.AppendLine("## Visual Language")
                builder.AppendLine(plan.VisualLanguage)
                builder.AppendLine()
                builder.AppendLine("## Photography Style")
                builder.AppendLine(plan.PhotographyStyle)
                builder.AppendLine()
                builder.AppendLine("## Layout Guidance")
                builder.AppendLine(plan.LayoutGuidance)
            Else
                builder.AppendLine("# 设计规划")
                builder.AppendLine()
                builder.AppendLine("## 产品概述")
                builder.AppendLine(plan.ProductSummary)
                builder.AppendLine()
                builder.AppendLine("## 设计主题")
                builder.AppendLine(plan.DesignTheme)
                builder.AppendLine()
                builder.AppendLine("## 目标人群")
                builder.AppendLine(plan.Audience)
                builder.AppendLine()
                builder.AppendLine("## 色彩系统")
                builder.AppendLine(plan.ColorSystem)
                builder.AppendLine()
                builder.AppendLine("## 字体系统")
                builder.AppendLine(plan.Typography)
                builder.AppendLine()
                builder.AppendLine("## 视觉语言")
                builder.AppendLine(plan.VisualLanguage)
                builder.AppendLine()
                builder.AppendLine("## 摄影风格")
                builder.AppendLine(plan.PhotographyStyle)
                builder.AppendLine()
                builder.AppendLine("## 版式指导")
                builder.AppendLine(plan.LayoutGuidance)
            End If

            Return builder.ToString().Trim()
        End Function

        Private Shared Function BuildImagePlanMarkdown(plans As IEnumerable(Of ImagePlanItem), targetLanguage As String, useEnglish As Boolean) As String
            Dim builder As New StringBuilder()
            Dim includeOverlayText = Not String.Equals(targetLanguage, NoTextLanguage, StringComparison.Ordinal)
            builder.AppendLine(If(useEnglish, "# Image Plan", "# 图片规划"))

            For Each plan In plans.OrderBy(Function(item) item.Sequence)
                builder.AppendLine()
                builder.AppendLine(If(useEnglish,
                                      $"## Image {plan.Sequence} - {plan.Title}",
                                      $"## 图{plan.Sequence} - {plan.Title}"))
                builder.AppendLine(If(useEnglish, $"- Purpose: {plan.Purpose}", $"- 用途: {plan.Purpose}"))
                builder.AppendLine(If(useEnglish, $"- Angle: {plan.Angle}", $"- 视角: {plan.Angle}"))
                builder.AppendLine(If(useEnglish, $"- Scene: {plan.Scene}", $"- 场景: {plan.Scene}"))
                builder.AppendLine(If(useEnglish, $"- Selling Point: {plan.SellingPoint}", $"- 卖点: {plan.SellingPoint}"))
                If includeOverlayText Then
                    If Not String.IsNullOrWhiteSpace(plan.OverlayTitle) Then
                        builder.AppendLine(If(useEnglish,
                                              $"- Title: ""{plan.OverlayTitle.Trim()}""",
                                              $"- 主标题: ""{plan.OverlayTitle.Trim()}"""))
                    End If
                    If Not String.IsNullOrWhiteSpace(plan.OverlaySubtitle) Then
                        builder.AppendLine(If(useEnglish,
                                              $"- Subtitle: ""{plan.OverlaySubtitle.Trim()}""",
                                              $"- 副标题: ""{plan.OverlaySubtitle.Trim()}"""))
                    End If
                End If
                builder.AppendLine(If(useEnglish, $"- Prompt: {plan.Prompt}", $"- 提示词: {plan.Prompt}"))
            Next

            Return builder.ToString().Trim()
        End Function

        Private Shared Function ParseImagePlanMarkdown(markdown As String) As List(Of ImagePlanItem)
            Dim plans As New List(Of ImagePlanItem)()
            Dim sections = Regex.Split(markdown, "(?m)^##\s+").Where(Function(item) Not String.IsNullOrWhiteSpace(item)).ToList()

            For Each section In sections
                Dim lines = section.Replace(vbCr, String.Empty).Split({vbLf}, StringSplitOptions.None).ToList()
                If lines.Count = 0 Then Continue For

                Dim header = lines(0).Trim()
                If header.StartsWith("# Image Plan", StringComparison.OrdinalIgnoreCase) OrElse header.StartsWith("Image Plan", StringComparison.OrdinalIgnoreCase) Then Continue For
                If Not Regex.IsMatch(header, "^(Image|图)\s*\d+", RegexOptions.IgnoreCase) Then Continue For

                Dim plan As New ImagePlanItem()
                Dim sequenceMatch = Regex.Match(header, "(Image|图)\s*(\d+)", RegexOptions.IgnoreCase)
                If sequenceMatch.Success Then Integer.TryParse(sequenceMatch.Groups(2).Value, plan.Sequence)

                Dim dashIndex = header.IndexOf("-"c)
                plan.Title = If(dashIndex >= 0 AndAlso dashIndex < header.Length - 1, header.Substring(dashIndex + 1).Trim(), header)

                For Each rawLine In lines.Skip(1)
                    Dim line = rawLine.Trim()
                    If String.IsNullOrWhiteSpace(line) OrElse Not line.StartsWith("-"c) Then Continue For

                    Dim colonIndex = line.IndexOf(":"c)
                    If colonIndex < 0 Then
                        colonIndex = line.IndexOf("："c)
                    End If
                    If colonIndex < 0 Then Continue For

                    Dim key = line.Substring(1, colonIndex - 1).Trim().ToLowerInvariant()
                    Dim value = line.Substring(colonIndex + 1).Trim()

                    Select Case key
                        Case "purpose", "用途"
                            plan.Purpose = value
                        Case "angle", "视角"
                            plan.Angle = value
                        Case "scene", "场景"
                            plan.Scene = value
                        Case "sellingpoint", "卖点"
                            plan.SellingPoint = value
                        Case "title", "主标题"
                            plan.OverlayTitle = UnquoteDisplayText(value)
                        Case "subtitle", "副标题"
                            plan.OverlaySubtitle = UnquoteDisplayText(value)
                        Case "prompt", "提示词"
                            plan.Prompt = value
                    End Select
                Next

                If plan.Sequence <= 0 Then plan.Sequence = plans.Count + 1
                If String.IsNullOrWhiteSpace(plan.Purpose) Then plan.Purpose = plan.Title
                If String.IsNullOrWhiteSpace(plan.Prompt) Then plan.Prompt = $"{plan.Title}, {plan.Scene}, {plan.SellingPoint}"
                plans.Add(plan)
            Next

            Return plans.OrderBy(Function(item) item.Sequence).ToList()
        End Function

        Private Shared Function ExpandImagePlans(basePlans As IEnumerable(Of ImagePlanItem), derivativeCount As Integer) As List(Of ImagePlanItem)
            Dim results As New List(Of ImagePlanItem)()
            Dim safeDerivativeCount = Math.Min(10, Math.Max(1, derivativeCount))
            Dim renderSequence = 1

            For Each basePlan In basePlans.OrderBy(Function(item) item.Sequence)
                For variantIndex = 1 To safeDerivativeCount
                    Dim expandedTitle = basePlan.Title
                    If safeDerivativeCount > 1 Then
                        expandedTitle = $"{basePlan.Title} {variantIndex}/{safeDerivativeCount}"
                    End If

                    results.Add(New ImagePlanItem With {
                        .Sequence = renderSequence,
                        .Title = expandedTitle,
                        .Purpose = basePlan.Purpose,
                        .Angle = basePlan.Angle,
                        .Scene = basePlan.Scene,
                        .SellingPoint = basePlan.SellingPoint,
                        .OverlayTitle = basePlan.OverlayTitle,
                        .OverlaySubtitle = basePlan.OverlaySubtitle,
                        .Prompt = basePlan.Prompt
                    })
                    renderSequence += 1
                Next
            Next

            Return results
        End Function

        Private Sub EnsureProjectModelSelections()
            If String.IsNullOrWhiteSpace(_selectedPlannerProviderKey) Then
                _selectedPlannerProviderKey = _vertexService.GetDefaultPlannerProviderKey()
            End If

            Dim plannerOptions = GetPlannerModelOptionsForProvider(_selectedPlannerProviderKey)
            If String.IsNullOrWhiteSpace(_selectedPlannerModel) OrElse
               String.Equals(_selectedPlannerModel, "Auto", StringComparison.OrdinalIgnoreCase) OrElse
               Not plannerOptions.Any(Function(item) String.Equals(item, _selectedPlannerModel, StringComparison.OrdinalIgnoreCase)) Then
                _selectedPlannerModel = _vertexService.GetDefaultPlannerModel(_selectedPlannerProviderKey)
            End If

            If String.IsNullOrWhiteSpace(_selectedImageProviderKey) Then
                _selectedImageProviderKey = _vertexService.GetDefaultImageProviderKey()
            End If

            Dim imageOptions = GetImageModelOptionsForProvider(_selectedImageProviderKey)
            If String.IsNullOrWhiteSpace(_selectedImageModel) OrElse
               Not imageOptions.Any(Function(item) String.Equals(item, _selectedImageModel, StringComparison.OrdinalIgnoreCase)) Then
                _selectedImageModel = _vertexService.GetDefaultImageModel(_selectedImageProviderKey)
            End If
        End Sub

        Private Shared Function BuildModelOptions(primaryModel As String, fallbackModels As IEnumerable(Of String)) As IReadOnlyList(Of String)
            Dim values As New List(Of String)

            If Not String.IsNullOrWhiteSpace(primaryModel) Then
                values.Add(primaryModel.Trim())
            End If

            For Each model In fallbackModels
                If String.IsNullOrWhiteSpace(model) Then
                    Continue For
                End If

                Dim normalized = model.Trim()
                If Not values.Any(Function(item) String.Equals(item, normalized, StringComparison.OrdinalIgnoreCase)) Then
                    values.Add(normalized)
                End If
            Next

            Return values
        End Function

        Private Function GetPlannerModelOptionsForProvider(providerKey As String) As IReadOnlyList(Of String)
            If String.Equals(providerKey, "qwen", StringComparison.OrdinalIgnoreCase) Then
                Return BuildModelOptions(_vertexService.GetDefaultPlannerModel("qwen"), {"qwen3.5-vl-plus", "qwen-vl-max-latest"})
            End If

            If String.Equals(providerKey, "openai", StringComparison.OrdinalIgnoreCase) Then
                Return BuildModelOptions(_vertexService.GetDefaultPlannerModel("openai"), {"gpt-4.1-mini", "gpt-4.1", "gpt-4o-mini"})
            End If

            Return BuildModelOptions(_vertexService.GetDefaultPlannerModel("deepseek"), {"deepseek-chat", "deepseek-reasoner"})
        End Function

        Private Function GetImageModelOptionsForProvider(providerKey As String) As IReadOnlyList(Of String)
            If String.Equals(providerKey, "qwen", StringComparison.OrdinalIgnoreCase) Then
                Return BuildModelOptions(_vertexService.GetDefaultImageModel("qwen"), {"qwen-image-2.0-pro"})
            End If

            If String.Equals(providerKey, "openai", StringComparison.OrdinalIgnoreCase) Then
                Return BuildModelOptions(_vertexService.GetDefaultImageModel("openai"), {"gpt-image-1.5", "gpt-image-1", "gpt-image-1-mini"})
            End If

            Return BuildModelOptions(_vertexService.GetDefaultImageModel("gemini"), {"gemini-3.1-flash-image-preview", "gemini-2.5-flash-image"})
        End Function

        Private Shared Function BuildFallbackProjectSummary(requirementText As String) As String
            Dim normalized = NormalizeSummaryText(requirementText)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return String.Empty
            End If

            Dim sentenceBreakIndex = normalized.IndexOfAny({".", "。", "!", "！", "?", "？", ";", "；", vbCr, vbLf}.Select(Function(ch) CChar(ch)).ToArray())
            If sentenceBreakIndex > 0 Then
                normalized = normalized.Substring(0, sentenceBreakIndex).Trim()
            End If

            If normalized.Length > AutoProjectSummaryMaxLength Then
                normalized = normalized.Substring(0, AutoProjectSummaryMaxLength).TrimEnd() & "..."
            End If

            Return normalized
        End Function

        Private Shared Function NormalizeSummaryText(value As String) As String
            Dim normalized = If(value, String.Empty).Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            normalized = Regex.Replace(normalized, "\s+", " ")
            Return normalized
        End Function

        Private Shared Function IsDefaultProjectSummary(value As String) As Boolean
            Dim normalized = NormalizeSummaryText(value)
            Return String.IsNullOrWhiteSpace(normalized) OrElse
                   String.Equals(normalized, DefaultProjectSummaryText, StringComparison.Ordinal)
        End Function

        Private Shared Function NormalizeAspectRatio(value As String) As String
            Return value.Trim()
        End Function

        Private Shared Function NormalizeCustomAspectRatio(value As String) As String
            Dim normalized = If(value, String.Empty).Trim().Replace("x", "*")
            Dim parts = normalized.Split("*"c)
            If parts.Length <> 2 Then
                Return "1:1"
            End If

            Dim width As Integer
            Dim height As Integer
            If Not Integer.TryParse(parts(0), width) OrElse Not Integer.TryParse(parts(1), height) OrElse width <= 0 OrElse height <= 0 Then
                Return "1:1"
            End If

            Return $"{width}:{height}"
        End Function

        Private Function NormalizeImageSize(value As String) As String
            Dim normalized = If(value, String.Empty).Trim()
            If _vertexService.UsesCustomImageSizeForProvider(SelectedImageProviderKey) Then
                Dim parts = normalized.Replace("x", "*").Split("*"c)
                If parts.Length <> 2 Then
                    Throw New InvalidOperationException("Qwen 绘图尺寸格式应为 宽*高，例如 1536*1024。")
                End If

                Dim width As Integer
                Dim height As Integer
                If Not Integer.TryParse(parts(0), width) OrElse Not Integer.TryParse(parts(1), height) Then
                    Throw New InvalidOperationException("Qwen 绘图尺寸必须是整数，格式示例：1536*1024。")
                End If

                Dim pixels = CLng(width) * CLng(height)
                If pixels < 512L * 512L OrElse pixels > 2048L * 2048L Then
                    Throw New InvalidOperationException("Qwen 绘图尺寸总像素必须在 512*512 到 2048*2048 之间。")
                End If

                Return $"{width}*{height}"
            End If

            Return normalized.ToUpperInvariant()
        End Function

        Private Function GetParsedCustomImageSize() As Tuple(Of Integer, Integer)
            Dim widthValue = 1024
            Dim heightValue = 1024
            Dim normalized = If(_outputResolution, String.Empty).Trim().Replace("x", "*")
            Dim parts = normalized.Split("*"c)
            If parts.Length = 2 Then
                Integer.TryParse(parts(0), widthValue)
                Integer.TryParse(parts(1), heightValue)
            End If

            Return Tuple.Create(Math.Max(1, widthValue), Math.Max(1, heightValue))
        End Function

        Private Sub UpdateCustomImageSize(width As Integer, height As Integer)
            Dim safeWidth = Math.Max(1, width)
            Dim safeHeight = Math.Max(1, height)
            OutputResolution = $"{safeWidth}*{safeHeight}"
        End Sub

        Private Shared Function UnquoteDisplayText(value As String) As String
            Dim trimmed = value.Trim()
            If trimmed.Length >= 2 Then
                If (trimmed.StartsWith("""", StringComparison.Ordinal) AndAlso trimmed.EndsWith("""", StringComparison.Ordinal)) OrElse
                   (trimmed.StartsWith("“", StringComparison.Ordinal) AndAlso trimmed.EndsWith("”", StringComparison.Ordinal)) Then
                    Return trimmed.Substring(1, trimmed.Length - 2).Trim()
                End If
            End If
            Return trimmed
        End Function

        Private Shared Function LoadBitmap(filePath As String) As BitmapImage
            Using stream = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Dim bitmap As New BitmapImage()
                bitmap.BeginInit()
                bitmap.CacheOption = BitmapCacheOption.OnLoad
                bitmap.StreamSource = stream
                bitmap.EndInit()
                bitmap.Freeze()
                Return bitmap
            End Using
        End Function

        Private Shared Function IsSupportedImageFile(filePath As String) As Boolean
            Dim extension = Path.GetExtension(filePath)
            Return Not String.IsNullOrWhiteSpace(extension) AndAlso {".png", ".jpg", ".jpeg", ".bmp", ".webp"}.Contains(extension, StringComparer.OrdinalIgnoreCase)
        End Function

        Private Shared Function SanitizeFileName(value As String) As String
            Dim invalid = Path.GetInvalidFileNameChars()
            Dim builder As New StringBuilder()
            For Each ch In value
                builder.Append(If(invalid.Contains(ch), "_"c, ch))
            Next
            Return builder.ToString().Trim().Trim("."c)
        End Function

        Private Shared Function GetSortableFileName(filePath As String) As String
            Return Path.GetFileName(filePath)
        End Function

        Private Shared Function CreateBrush(hex As String) As Brush
            Dim brush = DirectCast(New BrushConverter().ConvertFromString(hex), Brush)
            brush.Freeze()
            Return brush
        End Function
    End Class
End Namespace

Imports System
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Threading
Imports System.Windows
Imports System.Windows.Data
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Threading
Imports Microsoft.Win32
Imports MediaFactory.Infrastructure
Imports MediaFactory.Localization
Imports MediaFactory.Models
Imports MediaFactory.Services

Namespace ViewModels
    Public Class StudioShellViewModel
        Inherits ObservableObject
        Implements IDisposable

        Public Enum StudioPage
            ProjectHall
            ProjectWorkspace
            TaskQueue
            MediaLibrary
            SystemSettings
        End Enum

        Public Enum ProjectFilter
            All
            Planning
            Rendering
            Ready
            Failed
            Completed
        End Enum

        Public Enum MediaLibraryFilter
            All
            Source
            Generated
            Failed
        End Enum

        Private Shared ReadOnly PersistedProjectProperties As HashSet(Of String) = New HashSet(Of String)(StringComparer.Ordinal) From {
            NameOf(ProjectSessionViewModel.ProjectName),
            NameOf(ProjectSessionViewModel.ProjectSummary),
            NameOf(ProjectSessionViewModel.SelectedScenarioPresetId),
            NameOf(ProjectSessionViewModel.SelectedPlannerProviderKey),
            NameOf(ProjectSessionViewModel.RequirementText),
            NameOf(ProjectSessionViewModel.TargetLanguage),
            NameOf(ProjectSessionViewModel.SelectedPlannerModel),
            NameOf(ProjectSessionViewModel.SelectedImageProviderKey),
            NameOf(ProjectSessionViewModel.SelectedImageModel),
            NameOf(ProjectSessionViewModel.SelectedAspectRatio),
            NameOf(ProjectSessionViewModel.OutputResolution),
            NameOf(ProjectSessionViewModel.RequestedCount),
            NameOf(ProjectSessionViewModel.DerivativeImageCount),
            NameOf(ProjectSessionViewModel.AutoRenderAfterPlan),
            NameOf(ProjectSessionViewModel.DesignPlanMarkdown),
            NameOf(ProjectSessionViewModel.ImagePlanMarkdown),
            NameOf(ProjectSessionViewModel.HasGeneratedPlan),
            NameOf(ProjectSessionViewModel.LastOutputDirectory),
            NameOf(ProjectSessionViewModel.LastActivatedAt)
        }

        Private ReadOnly _vertexService As VertexAiService
        Private ReadOnly _persistenceService As StudioPersistenceService
        Private ReadOnly _projectHallSource As CollectionViewSource
        Private ReadOnly _taskQueueSource As CollectionViewSource
        Private ReadOnly _mediaLibrarySource As CollectionViewSource
        Private ReadOnly _taskExecutionSemaphore As SemaphoreSlim
        Private ReadOnly _plannerProviderEditors As ObservableCollection(Of ModelProviderEditor)
        Private ReadOnly _imageProviderEditors As ObservableCollection(Of ModelProviderEditor)
        Private ReadOnly _deepSeekPlannerModelOptionItems As ObservableCollection(Of ProviderOptionItem)
        Private ReadOnly _openAiPlannerModelOptionItems As ObservableCollection(Of ProviderOptionItem)
        Private ReadOnly _qwenPlannerModelOptionItems As ObservableCollection(Of ProviderOptionItem)
        Private ReadOnly _geminiImageModelOptionItems As ObservableCollection(Of ProviderOptionItem)
        Private ReadOnly _openAiImageModelOptionItems As ObservableCollection(Of ProviderOptionItem)
        Private ReadOnly _qwenImageModelOptionItems As ObservableCollection(Of ProviderOptionItem)
        Private _currentPage As StudioPage = StudioPage.ProjectHall
        Private _activeFilter As ProjectFilter = ProjectFilter.All
        Private _activeMediaFilter As MediaLibraryFilter = MediaLibraryFilter.All
        Private _selectedProject As ProjectSessionViewModel
        Private _selectedMediaLibraryItem As MediaLibraryItem
        Private _projectCounter As Integer
        Private _searchText As String = String.Empty
        Private _languageCode As String = "en"
        Private _selectedScenarioPreset As ScenarioPreset
        Private _modelSettings As ModelProviderConfiguration
        Private _isDisposed As Boolean
        Private _isRestoringState As Boolean
        Private _isRefreshDeferred As Boolean

        Public Sub New()
            _vertexService = New VertexAiService()
            _persistenceService = New StudioPersistenceService()
            _taskExecutionSemaphore = New SemaphoreSlim(5, 5)
            _modelSettings = VertexAiService.LoadModelConfiguration()
            _plannerProviderEditors = New ObservableCollection(Of ModelProviderEditor)()
            _imageProviderEditors = New ObservableCollection(Of ModelProviderEditor)()
            _deepSeekPlannerModelOptionItems = New ObservableCollection(Of ProviderOptionItem)()
            _openAiPlannerModelOptionItems = New ObservableCollection(Of ProviderOptionItem)()
            _qwenPlannerModelOptionItems = New ObservableCollection(Of ProviderOptionItem)()
            _geminiImageModelOptionItems = New ObservableCollection(Of ProviderOptionItem)()
            _openAiImageModelOptionItems = New ObservableCollection(Of ProviderOptionItem)()
            _qwenImageModelOptionItems = New ObservableCollection(Of ProviderOptionItem)()
            Projects = New ObservableCollection(Of ProjectSessionViewModel)()
            OpenWorkspaceProjects = New ObservableCollection(Of ProjectSessionViewModel)()
            ScenarioPresets = New ObservableCollection(Of ScenarioPreset)()
            MediaLibraryItems = New ObservableCollection(Of MediaLibraryItem)()
            AddHandler OpenWorkspaceProjects.CollectionChanged, AddressOf OpenWorkspaceProjects_CollectionChanged
            AddHandler ScenarioPresets.CollectionChanged, AddressOf ScenarioPresets_CollectionChanged

            _projectHallSource = New CollectionViewSource With {.Source = Projects}
            AddHandler _projectHallSource.Filter, AddressOf ProjectHallSource_Filter
            FilteredProjects = _projectHallSource.View
            FilteredProjects.SortDescriptions.Add(New SortDescription(NameOf(ProjectSessionViewModel.IsBusy), ListSortDirection.Descending))
            FilteredProjects.SortDescriptions.Add(New SortDescription(NameOf(ProjectSessionViewModel.LastActivatedAt), ListSortDirection.Descending))
            FilteredProjects.SortDescriptions.Add(New SortDescription(NameOf(ProjectSessionViewModel.ProjectName), ListSortDirection.Ascending))

            _taskQueueSource = New CollectionViewSource With {.Source = Projects}
            AddHandler _taskQueueSource.Filter, AddressOf TaskQueueSource_Filter
            FilteredTaskProjects = _taskQueueSource.View
            FilteredTaskProjects.SortDescriptions.Add(New SortDescription(NameOf(ProjectSessionViewModel.IsBusy), ListSortDirection.Descending))
            FilteredTaskProjects.SortDescriptions.Add(New SortDescription(NameOf(ProjectSessionViewModel.ProgressPercent), ListSortDirection.Descending))
            FilteredTaskProjects.SortDescriptions.Add(New SortDescription(NameOf(ProjectSessionViewModel.ProjectName), ListSortDirection.Ascending))

            _mediaLibrarySource = New CollectionViewSource With {.Source = MediaLibraryItems}
            AddHandler _mediaLibrarySource.Filter, AddressOf MediaLibrarySource_Filter
            FilteredMediaLibraryItems = _mediaLibrarySource.View
            FilteredMediaLibraryItems.SortDescriptions.Add(New SortDescription(NameOf(MediaLibraryItem.SortDate), ListSortDirection.Descending))
            FilteredMediaLibraryItems.SortDescriptions.Add(New SortDescription(NameOf(MediaLibraryItem.ProjectName), ListSortDirection.Ascending))

            RestoreScenarioPresets()
            RestorePersistedState()
            LocalizationManager.Instance.CurrentLanguageCode = _languageCode
            AttachModelSettings(_modelSettings)
            ReloadModelProviderEditors()
            RefreshModelOptionCollections()
            If Projects.Count = 0 Then
                AddNewProject(openWorkspace:=False)
            End If

            If SelectedProject Is Nothing Then
                SelectedProject = Projects.FirstOrDefault()
            End If

            If CurrentPage = StudioPage.ProjectWorkspace AndAlso SelectedProject IsNot Nothing Then
                EnsureWorkspaceTab(SelectedProject)
            ElseIf CurrentPage = StudioPage.ProjectWorkspace AndAlso OpenWorkspaceProjects.Count = 0 Then
                CurrentPage = StudioPage.ProjectHall
            End If

            RefreshFilteredCollections()
            NotifyPageStateChanged()
        End Sub

        Public ReadOnly Property Projects As ObservableCollection(Of ProjectSessionViewModel)
        Public ReadOnly Property OpenWorkspaceProjects As ObservableCollection(Of ProjectSessionViewModel)
        Public ReadOnly Property ScenarioPresets As ObservableCollection(Of ScenarioPreset)
        Public ReadOnly Property MediaLibraryItems As ObservableCollection(Of MediaLibraryItem)
        Public ReadOnly Property FilteredProjects As ICollectionView
        Public ReadOnly Property FilteredTaskProjects As ICollectionView
        Public ReadOnly Property FilteredMediaLibraryItems As ICollectionView

        Public Property SearchText As String
            Get
                Return _searchText
            End Get
            Set(value As String)
                If SetProperty(_searchText, value) Then
                    RefreshFilteredCollections()
                    PersistStudioState()
                End If
            End Set
        End Property

        Public Property LanguageCode As String
            Get
                Return _languageCode
            End Get
            Set(value As String)
                Dim normalized = LocalizationManager.NormalizeLanguageCode(value)
                If SetProperty(_languageCode, normalized) Then
                    LocalizationManager.Instance.CurrentLanguageCode = normalized
                    For Each project In Projects
                        project.NotifyLocalizationChanged()
                    Next
                    NotifyDashboardChanged()
                    NotifyPageStateChanged()
                    OnPropertyChanged(NameOf(WorkspaceTitle))
                    OnPropertyChanged(NameOf(WorkspaceSubtitle))
                    OnPropertyChanged(NameOf(StatusBarSummary))
                    PersistStudioState()
                End If
            End Set
        End Property

        Public Property SelectedScenarioPreset As ScenarioPreset
            Get
                Return _selectedScenarioPreset
            End Get
            Set(value As ScenarioPreset)
                If SetProperty(_selectedScenarioPreset, value) Then
                    OnPropertyChanged(NameOf(HasSelectedScenarioPreset))
                End If
            End Set
        End Property

        Public Property ModelSettings As ModelProviderConfiguration
            Get
                Return _modelSettings
            End Get
            Set(value As ModelProviderConfiguration)
                If ReferenceEquals(_modelSettings, value) Then
                    Return
                End If

                AttachModelSettings(Nothing)
                If SetProperty(_modelSettings, value) Then
                    AttachModelSettings(value)
                    ReloadModelProviderEditors()
                    RefreshModelOptionCollections()
                    NotifySystemSettingsModelBindingsChanged()
                End If
            End Set
        End Property

        Public ReadOnly Property LanguageOptions As IReadOnlyList(Of ProviderOptionItem) = New List(Of ProviderOptionItem) From {
            New ProviderOptionItem With {.Key = "en", .DisplayName = "English"},
            New ProviderOptionItem With {.Key = "zh-CN", .DisplayName = "简体中文"}
        }

        Public ReadOnly Property PlannerProviderEditors As ObservableCollection(Of ModelProviderEditor)
            Get
                Return _plannerProviderEditors
            End Get
        End Property

        Public ReadOnly Property ImageProviderEditors As ObservableCollection(Of ModelProviderEditor)
            Get
                Return _imageProviderEditors
            End Get
        End Property

        Public ReadOnly Property DeepSeekPlannerModelOptions As IReadOnlyList(Of String) = New List(Of String) From {
            "deepseek-chat", "deepseek-reasoner"
        }

        Public ReadOnly Property OpenAiPlannerModelOptions As IReadOnlyList(Of String) = New List(Of String) From {
            "gpt-4.1-mini", "gpt-4.1", "gpt-4o-mini"
        }

        Public ReadOnly Property QwenPlannerModelOptions As IReadOnlyList(Of String) = New List(Of String) From {
            "qwen3.5-vl-plus", "qwen-vl-max-latest"
        }

        Public ReadOnly Property GeminiImageModelOptions As IReadOnlyList(Of String) = New List(Of String) From {
            "gemini-3.1-flash-image-preview", "gemini-2.5-flash-image"
        }

        Public ReadOnly Property OpenAiImageModelOptions As IReadOnlyList(Of String) = New List(Of String) From {
            "gpt-image-1.5", "gpt-image-1", "gpt-image-1-mini"
        }

        Public ReadOnly Property QwenImageModelOptions As IReadOnlyList(Of String) = New List(Of String) From {
            "qwen-image-2.0-pro"
        }

        Public ReadOnly Property DeepSeekPlannerModelOptionItems As ObservableCollection(Of ProviderOptionItem)
            Get
                Return _deepSeekPlannerModelOptionItems
            End Get
        End Property

        Public ReadOnly Property QwenPlannerModelOptionItems As ObservableCollection(Of ProviderOptionItem)
            Get
                Return _qwenPlannerModelOptionItems
            End Get
        End Property

        Public ReadOnly Property OpenAiPlannerModelOptionItems As ObservableCollection(Of ProviderOptionItem)
            Get
                Return _openAiPlannerModelOptionItems
            End Get
        End Property

        Public ReadOnly Property GeminiImageModelOptionItems As ObservableCollection(Of ProviderOptionItem)
            Get
                Return _geminiImageModelOptionItems
            End Get
        End Property

        Public ReadOnly Property OpenAiImageModelOptionItems As ObservableCollection(Of ProviderOptionItem)
            Get
                Return _openAiImageModelOptionItems
            End Get
        End Property

        Public ReadOnly Property QwenImageModelOptionItems As ObservableCollection(Of ProviderOptionItem)
            Get
                Return _qwenImageModelOptionItems
            End Get
        End Property

        Public ReadOnly Property HasSelectedScenarioPreset As Boolean
            Get
                Return SelectedScenarioPreset IsNot Nothing
            End Get
        End Property

        Public Property SelectedProject As ProjectSessionViewModel
            Get
                Return _selectedProject
            End Get
            Set(value As ProjectSessionViewModel)
                If SetProperty(_selectedProject, value) Then
                    If value IsNot Nothing Then
                        value.MarkActivated()
                        PersistProject(value)
                    End If

                    OnPropertyChanged(NameOf(HasSelectedProject))
                    OnPropertyChanged(NameOf(WorkspaceTitle))
                    OnPropertyChanged(NameOf(WorkspaceSubtitle))
                    OnPropertyChanged(NameOf(CurrentPageDescription))
                    RefreshFilteredCollections()
                    PersistStudioState()
                End If
            End Set
        End Property

        Public Property SelectedMediaLibraryItem As MediaLibraryItem
            Get
                Return _selectedMediaLibraryItem
            End Get
            Set(value As MediaLibraryItem)
                If SetProperty(_selectedMediaLibraryItem, value) Then
                    OnPropertyChanged(NameOf(HasSelectedMediaLibraryItem))
                End If
            End Set
        End Property

        Public ReadOnly Property HasSelectedProject As Boolean
            Get
                Return SelectedProject IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property HasSelectedMediaLibraryItem As Boolean
            Get
                Return SelectedMediaLibraryItem IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property HasOpenWorkspaceProjects As Boolean
            Get
                Return OpenWorkspaceProjects.Count > 0
            End Get
        End Property

        Public ReadOnly Property WorkspaceTabsSummary As String
            Get
                Return F("Shell.WorkspaceTabsSummary", "{0} projects open", OpenWorkspaceProjects.Count)
            End Get
        End Property

        Public Property CurrentPageValue As Integer
            Get
                Return CInt(_currentPage)
            End Get
            Set(value As Integer)
                CurrentPage = CType(value, StudioPage)
            End Set
        End Property

        Public Property CurrentPage As StudioPage
            Get
                Return _currentPage
            End Get
            Set(value As StudioPage)
                If _currentPage <> value Then
                    _currentPage = value
                    If _currentPage = StudioPage.SystemSettings Then
                        NotifySystemSettingsModelBindingsChanged()
                    End If
                    NotifyPageStateChanged()
                    PersistStudioState()
                End If
            End Set
        End Property

        Public Property ActiveFilter As ProjectFilter
            Get
                Return _activeFilter
            End Get
            Set(value As ProjectFilter)
                If SetProperty(_activeFilter, value) Then
                    OnPropertyChanged(NameOf(IsAllFilter))
                    OnPropertyChanged(NameOf(IsPlanningFilter))
                    OnPropertyChanged(NameOf(IsRenderingFilter))
                    OnPropertyChanged(NameOf(IsReadyFilter))
                    OnPropertyChanged(NameOf(IsFailedFilter))
                    OnPropertyChanged(NameOf(IsCompletedFilter))
                    RefreshFilteredCollections()
                    PersistStudioState()
                End If
            End Set
        End Property

        Public Property ActiveMediaFilter As MediaLibraryFilter
            Get
                Return _activeMediaFilter
            End Get
            Set(value As MediaLibraryFilter)
                If SetProperty(_activeMediaFilter, value) Then
                    OnPropertyChanged(NameOf(IsAllMediaFilter))
                    OnPropertyChanged(NameOf(IsSourceMediaFilter))
                    OnPropertyChanged(NameOf(IsGeneratedMediaFilter))
                    OnPropertyChanged(NameOf(IsFailedMediaFilter))
                    RefreshFilteredCollections()
                End If
            End Set
        End Property

        Public ReadOnly Property IsAllFilter As Boolean
            Get
                Return ActiveFilter = ProjectFilter.All
            End Get
        End Property

        Public ReadOnly Property IsPlanningFilter As Boolean
            Get
                Return ActiveFilter = ProjectFilter.Planning
            End Get
        End Property

        Public ReadOnly Property IsRenderingFilter As Boolean
            Get
                Return ActiveFilter = ProjectFilter.Rendering
            End Get
        End Property

        Public ReadOnly Property IsReadyFilter As Boolean
            Get
                Return ActiveFilter = ProjectFilter.Ready
            End Get
        End Property

        Public ReadOnly Property IsFailedFilter As Boolean
            Get
                Return ActiveFilter = ProjectFilter.Failed
            End Get
        End Property

        Public ReadOnly Property IsCompletedFilter As Boolean
            Get
                Return ActiveFilter = ProjectFilter.Completed
            End Get
        End Property

        Public ReadOnly Property IsAllMediaFilter As Boolean
            Get
                Return ActiveMediaFilter = MediaLibraryFilter.All
            End Get
        End Property

        Public ReadOnly Property IsSourceMediaFilter As Boolean
            Get
                Return ActiveMediaFilter = MediaLibraryFilter.Source
            End Get
        End Property

        Public ReadOnly Property IsGeneratedMediaFilter As Boolean
            Get
                Return ActiveMediaFilter = MediaLibraryFilter.Generated
            End Get
        End Property

        Public ReadOnly Property IsFailedMediaFilter As Boolean
            Get
                Return ActiveMediaFilter = MediaLibraryFilter.Failed
            End Get
        End Property

        Public ReadOnly Property IsProjectHallPage As Boolean
            Get
                Return CurrentPage = StudioPage.ProjectHall
            End Get
        End Property

        Public ReadOnly Property IsWorkspacePage As Boolean
            Get
                Return CurrentPage = StudioPage.ProjectWorkspace
            End Get
        End Property

        Public ReadOnly Property IsTaskQueuePage As Boolean
            Get
                Return CurrentPage = StudioPage.TaskQueue
            End Get
        End Property

        Public ReadOnly Property IsMediaLibraryPage As Boolean
            Get
                Return CurrentPage = StudioPage.MediaLibrary
            End Get
        End Property

        Public ReadOnly Property IsSystemSettingsPage As Boolean
            Get
                Return CurrentPage = StudioPage.SystemSettings
            End Get
        End Property

        Public ReadOnly Property CurrentPageTitle As String
            Get
                Select Case CurrentPage
                    Case StudioPage.ProjectWorkspace
                        Return WorkspaceTitle
                    Case StudioPage.TaskQueue
                        Return T("Shell.Page.TaskQueue", "Task Queue")
                    Case StudioPage.MediaLibrary
                        Return T("Shell.Page.MediaLibrary", "Media Library")
                    Case StudioPage.SystemSettings
                        Return T("Shell.Page.SystemSettings", "System Settings")
                    Case Else
                        Return T("Shell.Page.ProjectHall", "Project Hall")
                End Select
            End Get
        End Property

        Public ReadOnly Property CurrentPageDescription As String
            Get
                Select Case CurrentPage
                    Case StudioPage.ProjectWorkspace
                        If Not HasOpenWorkspaceProjects Then
                            Return T("Shell.Desc.WorkspaceEmpty", "Open a project from Project Hall or Task Queue to work with it here in tabs.")
                        End If

                        Return F("Shell.Desc.WorkspaceOpen", "{0} tabs open. Currently viewing {1}.", OpenWorkspaceProjects.Count, WorkspaceTitle)
                    Case StudioPage.TaskQueue
                        Return T("Shell.Desc.TaskQueue", "Track planning, rendering, and completion states for every project in one queue.")
                    Case StudioPage.MediaLibrary
                        Return F("Shell.Desc.MediaLibrary", "{0} media items collected. Filter quickly by type, status, and keyword.", TotalMediaCount)
                    Case StudioPage.SystemSettings
                        Return T("Shell.Desc.SystemSettings", "Manage providers, scenario templates, language, and default generation behavior.")
                    Case Else
                        Return F("Shell.Desc.ProjectHall", "{0} projects available. Planning and rendering can run in parallel.", Projects.Count)
                End Select
            End Get
        End Property

        Public ReadOnly Property WorkspaceTitle As String
            Get
                If SelectedProject Is Nothing OrElse Not HasOpenWorkspaceProjects Then
                    Return T("Shell.WorkspaceDefaultTitle", "Project Workspace")
                End If

                Return SelectedProject.ProjectName
            End Get
        End Property

        Public ReadOnly Property WorkspaceSubtitle As String
            Get
                If SelectedProject Is Nothing OrElse Not HasOpenWorkspaceProjects Then
                    Return T("Shell.WorkspaceDefaultSubtitle", "Open projects from Project Hall or Task Queue to switch between them here.")
                End If

                Return SelectedProject.ProjectSubtitle
            End Get
        End Property

        Public ReadOnly Property StatusBarSummary As String
            Get
                Return F("Shell.StatusBar", "Active projects {0} | Rendering {1} | Average progress {2}", ActiveProjectCount, RenderingProjectCount, AverageProgressText)
            End Get
        End Property

        Public ReadOnly Property ActiveProjectCount As Integer
            Get
                Return Projects.Where(Function(project) Not project.IsArchived AndAlso (project.UploadedImages.Count > 0 OrElse project.HasGeneratedPlan OrElse project.GeneratedImages.Count > 0 OrElse project.IsBusy)).Count()
            End Get
        End Property

        Public ReadOnly Property PlanningProjectCount As Integer
            Get
                Return Projects.Where(Function(project) Not project.IsArchived AndAlso project.IsPlanning).Count()
            End Get
        End Property

        Public ReadOnly Property RenderingProjectCount As Integer
            Get
                Return Projects.Where(Function(project) Not project.IsArchived AndAlso project.IsGeneratingImages).Count()
            End Get
        End Property

        Public ReadOnly Property CompletedProjectCount As Integer
            Get
                Return Projects.Where(Function(project) Not project.IsArchived AndAlso project.GeneratedImages.Any(Function(item) String.Equals(item.StatusText, "Success", StringComparison.Ordinal))).Count()
            End Get
        End Property

        Public ReadOnly Property FailedProjectCount As Integer
            Get
                Return Projects.Where(Function(project) Not project.IsArchived AndAlso project.HasRetryableFailures).Count()
            End Get
        End Property

        Public ReadOnly Property TotalMediaCount As Integer
            Get
                Return MediaLibraryItems.Count
            End Get
        End Property

        Public ReadOnly Property SourceMediaCount As Integer
            Get
                Return MediaLibraryItems.Where(Function(item) item.IsSourceImage).Count()
            End Get
        End Property

        Public ReadOnly Property GeneratedMediaCount As Integer
            Get
                Return MediaLibraryItems.Where(Function(item) item.IsGeneratedImage).Count()
            End Get
        End Property

        Public ReadOnly Property FailedMediaCount As Integer
            Get
                Return MediaLibraryItems.Where(Function(item) item.IsFailed).Count()
            End Get
        End Property

        Public ReadOnly Property HasFilteredProjects As Boolean
            Get
                Return FilteredProjects IsNot Nothing AndAlso Not FilteredProjects.IsEmpty
            End Get
        End Property

        Public ReadOnly Property HasFilteredTaskProjects As Boolean
            Get
                Return FilteredTaskProjects IsNot Nothing AndAlso Not FilteredTaskProjects.IsEmpty
            End Get
        End Property

        Public ReadOnly Property HasFilteredMediaLibraryItems As Boolean
            Get
                Return FilteredMediaLibraryItems IsNot Nothing AndAlso Not FilteredMediaLibraryItems.IsEmpty
            End Get
        End Property

        Public ReadOnly Property FilteredMediaCountText As String
            Get
                Dim count = If(FilteredMediaLibraryItems Is Nothing, 0, FilteredMediaLibraryItems.Cast(Of Object)().Count())
                Return $"当前显示 {count} 张"
            End Get
        End Property

        Public ReadOnly Property FilteredRetryableProjectCount As Integer
            Get
                If FilteredTaskProjects Is Nothing Then Return 0
                Return FilteredTaskProjects.Cast(Of Object)().
                    OfType(Of ProjectSessionViewModel)().
                    Count(Function(project) project.HasRetryableFailures)
            End Get
        End Property

        Public ReadOnly Property CanRetryFilteredFailures As Boolean
            Get
                Return FilteredRetryableProjectCount > 0
            End Get
        End Property

        Public ReadOnly Property StartSelectedQueueProjectsButtonText As String
            Get
                Return F("Queue.Action.StartSelectedCount", "Run Selected ({0})", SelectedQueueProjectCount)
            End Get
        End Property

        Public ReadOnly Property RetryFilteredFailuresButtonText As String
            Get
                Return F("Queue.Action.RetryFilteredCount", "Retry Failed ({0})", FilteredRetryableProjectCount)
            End Get
        End Property

        Public ReadOnly Property SelectedQueueProjectCount As Integer
            Get
                Return Projects.Where(Function(project) project.IsQueueSelected).Count()
            End Get
        End Property

        Public ReadOnly Property CanStartSelectedQueueProjects As Boolean
            Get
                Return Projects.Any(Function(project) project.IsQueueSelected AndAlso CanStartProjectTask(project))
            End Get
        End Property

        Public Property IsAllFilteredQueueProjectsSelected As Boolean
            Get
                Dim filteredProjects = GetFilteredTaskProjects().ToList()
                If filteredProjects.Count = 0 Then
                    Return False
                End If

                Return filteredProjects.All(Function(project) project.IsQueueSelected)
            End Get
            Set(value As Boolean)
                For Each project In GetFilteredTaskProjects()
                    project.IsQueueSelected = value
                Next

                RefreshFilteredCollections()
            End Set
        End Property

        Public ReadOnly Property AverageProgressValue As Integer
            Get
                If Projects.Count = 0 Then
                    Return 0
                End If

                Return CInt(Math.Round(Projects.Average(Function(project) CDbl(project.ProgressPercent))))
            End Get
        End Property

        Public ReadOnly Property AverageProgressText As String
            Get
                Return $"{AverageProgressValue}%"
            End Get
        End Property

        Private Sub RestorePersistedState()
            _isRestoringState = True

            Try
                Dim records = _persistenceService.LoadProjects()
                For Each record In records
                    Dim project = New ProjectSessionViewModel(record.ProjectName, _vertexService, record.ProjectId)
                    project.BindScenarioPresets(ScenarioPresets)
                    project.RestoreFromRecord(record)
                    AttachProject(project)
                    Projects.Add(project)
                    UpdateProjectCounterFromName(project.ProjectName)
                Next

                Dim state = _persistenceService.LoadStudioState()
                _searchText = state.SearchText
                _activeFilter = CType(Math.Max(0, Math.Min([Enum].GetValues(GetType(ProjectFilter)).Length - 1, state.ActiveFilter)), ProjectFilter)
                _currentPage = CType(Math.Max(0, Math.Min([Enum].GetValues(GetType(StudioPage)).Length - 1, state.CurrentPage)), StudioPage)
                _selectedProject = Projects.FirstOrDefault(Function(project) String.Equals(project.ProjectId, state.SelectedProjectId, StringComparison.Ordinal))
                _languageCode = LocalizationManager.NormalizeLanguageCode(state.LanguageCode)
            Finally
                _isRestoringState = False
            End Try
        End Sub

        Private Sub RestoreScenarioPresets()
            ScenarioPresets.Clear()

            For Each record In _persistenceService.LoadScenarioPresets().OrderBy(Function(item) item.SortOrder).ThenBy(Function(item) item.Name)
                Dim preset = New ScenarioPreset With {
                    .Id = record.Id,
                    .Name = record.Name,
                    .Description = record.Description,
                    .DesignPlanningTemplate = record.DesignPlanningTemplate,
                    .ImagePlanningTemplate = record.ImagePlanningTemplate,
                    .IsBuiltIn = record.IsBuiltIn,
                    .SortOrder = record.SortOrder
                }
                ScenarioPresets.Add(preset)
            Next

            SelectedScenarioPreset = ScenarioPresets.FirstOrDefault()
        End Sub

        Private Sub UpdateProjectCounterFromName(projectName As String)
            If String.IsNullOrWhiteSpace(projectName) Then Return

            Dim digits = New String(projectName.Where(AddressOf Char.IsDigit).ToArray())
            Dim parsed As Integer
            If Integer.TryParse(digits, parsed) Then
                _projectCounter = Math.Max(_projectCounter, parsed)
            End If
        End Sub

        Private Function NextProjectName() As String
            _projectCounter += 1
            Return F("Project.DefaultName", "Project {0}", _projectCounter)
        End Function

        Private Sub AddNewProject(Optional openWorkspace As Boolean = True)
            Dim project = New ProjectSessionViewModel(NextProjectName(), _vertexService)
            project.BindScenarioPresets(ScenarioPresets)
            AttachProject(project)
            Projects.Add(project)
            PersistProject(project)

            If openWorkspace Then
                EnsureWorkspaceTab(project)
                SelectedProject = project
                CurrentPage = StudioPage.ProjectWorkspace
            ElseIf SelectedProject Is Nothing Then
                SelectedProject = project
            End If

            RefreshFilteredCollections()
        End Sub

        Private Sub OpenWorkspaceProjects_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs)
            OnPropertyChanged(NameOf(HasOpenWorkspaceProjects))
            OnPropertyChanged(NameOf(WorkspaceTabsSummary))
            OnPropertyChanged(NameOf(CurrentPageDescription))
            OnPropertyChanged(NameOf(WorkspaceTitle))
            OnPropertyChanged(NameOf(WorkspaceSubtitle))
        End Sub

        Private Sub EnsureWorkspaceTab(project As ProjectSessionViewModel)
            If project Is Nothing Then Return

            If Not OpenWorkspaceProjects.Contains(project) Then
                OpenWorkspaceProjects.Add(project)
            End If
        End Sub

        Private Sub AttachScenarioPreset(preset As ScenarioPreset)
            AddHandler preset.PropertyChanged, AddressOf ScenarioPreset_PropertyChanged
        End Sub

        Private Sub DetachScenarioPreset(preset As ScenarioPreset)
            RemoveHandler preset.PropertyChanged, AddressOf ScenarioPreset_PropertyChanged
        End Sub

        Private Sub ScenarioPresets_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs)
            If e.OldItems IsNot Nothing Then
                For Each preset As ScenarioPreset In e.OldItems
                    DetachScenarioPreset(preset)
                Next
            End If

            If e.NewItems IsNot Nothing Then
                For Each preset As ScenarioPreset In e.NewItems
                    AttachScenarioPreset(preset)
                Next
            End If

            For Each project In Projects
                project.BindScenarioPresets(ScenarioPresets)
                PersistProject(project)
            Next

            If SelectedScenarioPreset Is Nothing OrElse Not ScenarioPresets.Contains(SelectedScenarioPreset) Then
                SelectedScenarioPreset = ScenarioPresets.FirstOrDefault()
            End If

            RefreshFilteredCollections()
        End Sub

        Private Sub ScenarioPreset_PropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            Dim preset = TryCast(sender, ScenarioPreset)
            If preset Is Nothing Then Return

            SaveScenarioPreset(preset)

            For Each project In Projects
                project.NotifyScenarioPresetsChanged()
            Next

            RefreshFilteredCollections()
        End Sub

        Private Sub AttachProject(project As ProjectSessionViewModel)
            AddHandler project.PropertyChanged, AddressOf Project_PropertyChanged
            AddHandler project.UploadedImages.CollectionChanged, AddressOf UploadedImages_CollectionChanged
            AddHandler project.GeneratedImages.CollectionChanged, AddressOf GeneratedImages_CollectionChanged

            For Each image In project.GeneratedImages
                AddHandler image.PropertyChanged, AddressOf GeneratedImage_PropertyChanged
            Next
        End Sub

        Private Sub DetachProject(project As ProjectSessionViewModel)
            RemoveHandler project.PropertyChanged, AddressOf Project_PropertyChanged
            RemoveHandler project.UploadedImages.CollectionChanged, AddressOf UploadedImages_CollectionChanged
            RemoveHandler project.GeneratedImages.CollectionChanged, AddressOf GeneratedImages_CollectionChanged

            For Each image In project.GeneratedImages
                RemoveHandler image.PropertyChanged, AddressOf GeneratedImage_PropertyChanged
            Next
        End Sub

        Private Sub UploadedImages_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs)
            PersistFromCollection(sender)
            RefreshFilteredCollections()
        End Sub

        Private Sub GeneratedImages_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs)
            If e.OldItems IsNot Nothing Then
                For Each item As GeneratedImageItem In e.OldItems
                    RemoveHandler item.PropertyChanged, AddressOf GeneratedImage_PropertyChanged
                Next
            End If

            If e.NewItems IsNot Nothing Then
                For Each item As GeneratedImageItem In e.NewItems
                    AddHandler item.PropertyChanged, AddressOf GeneratedImage_PropertyChanged
                Next
            End If

            PersistFromCollection(sender)
            RefreshFilteredCollections()
        End Sub

        Private Sub GeneratedImage_PropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            Dim image = TryCast(sender, GeneratedImageItem)
            If image Is Nothing Then Return

            Dim project = Projects.FirstOrDefault(Function(entry) entry.GeneratedImages.Contains(image))
            PersistProject(project)
            RefreshFilteredCollections()
        End Sub

        Private Sub PersistFromCollection(sender As Object)
            Dim project = Projects.FirstOrDefault(Function(entry) ReferenceEquals(entry.UploadedImages, sender) OrElse ReferenceEquals(entry.GeneratedImages, sender))
            PersistProject(project)
        End Sub

        Private Sub ProjectHallSource_Filter(sender As Object, e As FilterEventArgs)
            Dim project = TryCast(e.Item, ProjectSessionViewModel)
            e.Accepted = project IsNot Nothing AndAlso Not project.IsArchived AndAlso MatchesSearch(project, SearchText.Trim()) AndAlso MatchesFilter(project)
        End Sub

        Private Sub TaskQueueSource_Filter(sender As Object, e As FilterEventArgs)
            Dim project = TryCast(e.Item, ProjectSessionViewModel)
            e.Accepted = project IsNot Nothing AndAlso Not project.IsArchived AndAlso MatchesSearch(project, SearchText.Trim()) AndAlso MatchesFilter(project)
        End Sub

        Private Sub MediaLibrarySource_Filter(sender As Object, e As FilterEventArgs)
            Dim item = TryCast(e.Item, MediaLibraryItem)
            e.Accepted = item IsNot Nothing AndAlso MatchesMediaLibrarySearch(item, SearchText.Trim()) AndAlso MatchesMediaLibraryFilter(item)
        End Sub

        Private Shared Function MatchesSearch(project As ProjectSessionViewModel, keyword As String) As Boolean
            If String.IsNullOrWhiteSpace(keyword) Then Return True

            Return project.ProjectName.IndexOf(keyword, StringComparison.CurrentCultureIgnoreCase) >= 0 OrElse
                   project.ProjectSummary.IndexOf(keyword, StringComparison.CurrentCultureIgnoreCase) >= 0 OrElse
                   project.SelectedScenarioPresetName.IndexOf(keyword, StringComparison.CurrentCultureIgnoreCase) >= 0 OrElse
                   project.RequirementText.IndexOf(keyword, StringComparison.CurrentCultureIgnoreCase) >= 0 OrElse
                   project.LocalizedStageText.IndexOf(keyword, StringComparison.CurrentCultureIgnoreCase) >= 0
        End Function

        Private Function MatchesFilter(project As ProjectSessionViewModel) As Boolean
            Select Case ActiveFilter
                Case ProjectFilter.Planning
                    Return project.IsPlanning
                Case ProjectFilter.Rendering
                    Return project.IsGeneratingImages
                Case ProjectFilter.Ready
                    Return project.HasGeneratedPlan AndAlso Not project.IsPlanning AndAlso Not project.IsGeneratingImages
                Case ProjectFilter.Failed
                    Return project.HasRetryableFailures
                Case ProjectFilter.Completed
                    Return project.GeneratedImages.Any(Function(item) String.Equals(item.StatusText, "Success", StringComparison.Ordinal))
                Case Else
                    Return True
            End Select
        End Function

        Private Shared Function MatchesMediaLibrarySearch(item As MediaLibraryItem, keyword As String) As Boolean
            If String.IsNullOrWhiteSpace(keyword) Then Return True

            Return item.Title.IndexOf(keyword, StringComparison.CurrentCultureIgnoreCase) >= 0 OrElse
                   item.Subtitle.IndexOf(keyword, StringComparison.CurrentCultureIgnoreCase) >= 0 OrElse
                   item.ProjectName.IndexOf(keyword, StringComparison.CurrentCultureIgnoreCase) >= 0 OrElse
                   item.KindText.IndexOf(keyword, StringComparison.CurrentCultureIgnoreCase) >= 0 OrElse
                   item.StatusText.IndexOf(keyword, StringComparison.CurrentCultureIgnoreCase) >= 0 OrElse
                   item.FileName.IndexOf(keyword, StringComparison.CurrentCultureIgnoreCase) >= 0
        End Function

        Private Function MatchesMediaLibraryFilter(item As MediaLibraryItem) As Boolean
            Select Case ActiveMediaFilter
                Case MediaLibraryFilter.Source
                    Return item.IsSourceImage
                Case MediaLibraryFilter.Generated
                    Return item.IsGeneratedImage
                Case MediaLibraryFilter.Failed
                    Return item.IsFailed
                Case Else
                    Return True
            End Select
        End Function

        Private Shared Function BuildModelOptionItems(currentValue As String, defaults As IEnumerable(Of String)) As IReadOnlyList(Of ProviderOptionItem)
            Dim items As New List(Of ProviderOptionItem)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each item In defaults.Where(Function(value) Not String.IsNullOrWhiteSpace(value))
                Dim normalized = item.Trim()
                If seen.Add(normalized) Then
                    items.Add(New ProviderOptionItem With {
                        .Key = normalized,
                        .DisplayName = normalized
                    })
                End If
            Next

            Dim normalizedCurrent = If(currentValue, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(normalizedCurrent) AndAlso seen.Add(normalizedCurrent) Then
                items.Add(New ProviderOptionItem With {
                    .Key = normalizedCurrent,
                    .DisplayName = normalizedCurrent
                })
            End If

            Return items
        End Function

        Private Sub AttachModelSettings(settings As ModelProviderConfiguration)
            If _modelSettings IsNot Nothing Then
                RemoveHandler _modelSettings.PropertyChanged, AddressOf ModelSettings_PropertyChanged
            End If

            If settings IsNot Nothing Then
                AddHandler settings.PropertyChanged, AddressOf ModelSettings_PropertyChanged
            End If
        End Sub

        Private Sub ModelSettings_PropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            Select Case e.PropertyName
                Case NameOf(ModelProviderConfiguration.DeepSeekModel),
                     NameOf(ModelProviderConfiguration.OpenAiPlannerModel),
                     NameOf(ModelProviderConfiguration.QwenPlannerModel),
                     NameOf(ModelProviderConfiguration.GeminiImageModel),
                     NameOf(ModelProviderConfiguration.OpenAiImageModel),
                     NameOf(ModelProviderConfiguration.QwenImageModel)
                    RefreshModelOptionCollections()
                    NotifySystemSettingsModelBindingsChanged()
            End Select
        End Sub

        Private Sub ReloadModelProviderEditors()
            Dim editorState = VertexAiService.LoadModelProviderEditors()
            ReplaceProviderEditors(_plannerProviderEditors, editorState.PlannerProviders)
            ReplaceProviderEditors(_imageProviderEditors, editorState.ImageProviders)
        End Sub

        Private Sub RefreshModelOptionCollections()
            ReplaceProviderOptions(_deepSeekPlannerModelOptionItems, BuildModelOptionItems(ModelSettings?.DeepSeekModel, DeepSeekPlannerModelOptions))
            ReplaceProviderOptions(_openAiPlannerModelOptionItems, BuildModelOptionItems(ModelSettings?.OpenAiPlannerModel, OpenAiPlannerModelOptions))
            ReplaceProviderOptions(_qwenPlannerModelOptionItems, BuildModelOptionItems(ModelSettings?.QwenPlannerModel, QwenPlannerModelOptions))
            ReplaceProviderOptions(_geminiImageModelOptionItems, BuildModelOptionItems(ModelSettings?.GeminiImageModel, GeminiImageModelOptions))
            ReplaceProviderOptions(_openAiImageModelOptionItems, BuildModelOptionItems(ModelSettings?.OpenAiImageModel, OpenAiImageModelOptions))
            ReplaceProviderOptions(_qwenImageModelOptionItems, BuildModelOptionItems(ModelSettings?.QwenImageModel, QwenImageModelOptions))
        End Sub

        Private Shared Sub ReplaceProviderOptions(target As ObservableCollection(Of ProviderOptionItem), items As IEnumerable(Of ProviderOptionItem))
            Dim nextItems = items.ToList()
            Dim sharedCount = Math.Min(target.Count, nextItems.Count)

            For index = 0 To sharedCount - 1
                Dim current = target(index)
                Dim [next] = nextItems(index)

                If Not String.Equals(current.Key, [next].Key, StringComparison.OrdinalIgnoreCase) OrElse
                   Not String.Equals(current.DisplayName, [next].DisplayName, StringComparison.Ordinal) Then
                    target(index) = New ProviderOptionItem With {
                        .Key = [next].Key,
                        .DisplayName = [next].DisplayName
                    }
                End If
            Next

            For index = target.Count - 1 To nextItems.Count Step -1
                target.RemoveAt(index)
            Next

            For index = target.Count To nextItems.Count - 1
                Dim item = nextItems(index)
                target.Add(New ProviderOptionItem With {
                    .Key = item.Key,
                    .DisplayName = item.DisplayName
                })
            Next
        End Sub

        Private Shared Sub ReplaceProviderEditors(target As ObservableCollection(Of ModelProviderEditor),
                                                  items As IEnumerable(Of ModelProviderEditor))
            target.Clear()
            For Each item In items
                target.Add(item)
            Next
        End Sub

        Private Sub RefreshFilteredCollections()
            Try
                RebuildMediaLibrary()
                FilteredProjects.Refresh()
                FilteredTaskProjects.Refresh()
                FilteredMediaLibraryItems.Refresh()
                OnPropertyChanged(NameOf(HasFilteredProjects))
                OnPropertyChanged(NameOf(HasFilteredTaskProjects))
                OnPropertyChanged(NameOf(HasFilteredMediaLibraryItems))
                OnPropertyChanged(NameOf(FilteredRetryableProjectCount))
                OnPropertyChanged(NameOf(CanRetryFilteredFailures))
                OnPropertyChanged(NameOf(RetryFilteredFailuresButtonText))
                OnPropertyChanged(NameOf(SelectedQueueProjectCount))
                OnPropertyChanged(NameOf(CanStartSelectedQueueProjects))
                OnPropertyChanged(NameOf(StartSelectedQueueProjectsButtonText))
                OnPropertyChanged(NameOf(IsAllFilteredQueueProjectsSelected))
                OnPropertyChanged(NameOf(FilteredMediaCountText))
                NotifyDashboardChanged()
            Catch ex As InvalidOperationException
                QueueDeferredRefresh()
            End Try
        End Sub

        Private Sub QueueDeferredRefresh()
            If _isRefreshDeferred OrElse _isDisposed Then Return
            _isRefreshDeferred = True

            Application.Current.Dispatcher.BeginInvoke(
                Sub()
                    _isRefreshDeferred = False
                    If _isDisposed Then Return

                    Try
                        RebuildMediaLibrary()
                        FilteredProjects.Refresh()
                        FilteredTaskProjects.Refresh()
                        FilteredMediaLibraryItems.Refresh()
                    Catch
                    End Try

                    OnPropertyChanged(NameOf(HasFilteredProjects))
                    OnPropertyChanged(NameOf(HasFilteredTaskProjects))
                    OnPropertyChanged(NameOf(HasFilteredMediaLibraryItems))
                    OnPropertyChanged(NameOf(FilteredRetryableProjectCount))
                    OnPropertyChanged(NameOf(CanRetryFilteredFailures))
                    OnPropertyChanged(NameOf(RetryFilteredFailuresButtonText))
                    OnPropertyChanged(NameOf(SelectedQueueProjectCount))
                    OnPropertyChanged(NameOf(CanStartSelectedQueueProjects))
                    OnPropertyChanged(NameOf(StartSelectedQueueProjectsButtonText))
                    OnPropertyChanged(NameOf(IsAllFilteredQueueProjectsSelected))
                    OnPropertyChanged(NameOf(FilteredMediaCountText))
                    NotifyDashboardChanged()
                End Sub,
                DispatcherPriority.Background)
        End Sub

        Private Function GetFilteredTaskProjects() As IEnumerable(Of ProjectSessionViewModel)
            If FilteredTaskProjects Is Nothing Then
                Return Enumerable.Empty(Of ProjectSessionViewModel)()
            End If

            Return FilteredTaskProjects.Cast(Of Object)().
                OfType(Of ProjectSessionViewModel)()
        End Function

        Private Sub RebuildMediaLibrary()
            Dim previousSelectionPath = SelectedMediaLibraryItem?.FilePath
            Dim items = New List(Of MediaLibraryItem)()

            For Each project In Projects
                items.AddRange(BuildProjectMediaItems(project))
            Next

            MediaLibraryItems.Clear()
            For Each item In items.OrderByDescending(Function(entry) entry.SortDate).ThenBy(Function(entry) entry.ProjectName).ThenBy(Function(entry) entry.Title)
                MediaLibraryItems.Add(item)
            Next

            SelectedMediaLibraryItem = MediaLibraryItems.FirstOrDefault(
                Function(item) Not String.IsNullOrWhiteSpace(previousSelectionPath) AndAlso
                               String.Equals(item.FilePath, previousSelectionPath, StringComparison.OrdinalIgnoreCase))

            If SelectedMediaLibraryItem Is Nothing Then
                SelectedMediaLibraryItem = MediaLibraryItems.FirstOrDefault()
            End If
        End Sub

        Private Shared Function BuildProjectMediaItems(project As ProjectSessionViewModel) As IEnumerable(Of MediaLibraryItem)
            Dim items As New List(Of MediaLibraryItem)()

            For Each sourceImage In project.UploadedImages
                Dim title = If(String.IsNullOrWhiteSpace(sourceImage.FileName),
                               T("Media.Title.UnnamedSource", "Untitled Source Image"),
                               Path.GetFileNameWithoutExtension(sourceImage.FileName))

                items.Add(New MediaLibraryItem With {
                    .Title = title,
                    .Subtitle = sourceImage.FileName,
                    .ProjectName = project.ProjectName,
                    .KindText = T("Media.Kind.Source", "Source"),
                    .StatusText = T("Media.Status.Imported", "Imported"),
                    .FilePath = sourceImage.FullPath,
                    .Preview = sourceImage.Preview,
                    .StatusBrush = CreateBrush("#2563EB"),
                    .SortDate = GetSortDate(sourceImage.FullPath, project.LastActivatedAt),
                    .IsSourceImage = True,
                    .IsGeneratedImage = False,
                    .IsFailed = False,
                    .OwnerProject = project
                })
            Next

            For Each generatedImage In project.GeneratedImages
                items.Add(New MediaLibraryItem With {
                    .Title = If(String.IsNullOrWhiteSpace(generatedImage.Title), T("Media.Title.UnnamedGenerated", "Untitled Generated Image"), generatedImage.Title),
                    .Subtitle = If(String.IsNullOrWhiteSpace(generatedImage.Description), Path.GetFileName(generatedImage.SavedPath), generatedImage.Description),
                    .ProjectName = project.ProjectName,
                    .KindText = T("Media.Kind.Generated", "Generated"),
                    .StatusText = LocalizeMediaStatus(generatedImage.StatusText),
                    .FilePath = generatedImage.SavedPath,
                    .Preview = generatedImage.Preview,
                    .StatusBrush = generatedImage.StatusBrush,
                    .SortDate = GetSortDate(generatedImage.SavedPath, project.LastActivatedAt),
                    .IsSourceImage = False,
                    .IsGeneratedImage = True,
                    .IsFailed = String.Equals(generatedImage.StatusText, "Failed", StringComparison.Ordinal) OrElse
                                String.Equals(generatedImage.StatusText, "Missing", StringComparison.Ordinal),
                    .OwnerProject = project
                })
            Next

            Return items
        End Function

        Private Shared Function GetSortDate(filePath As String, fallback As DateTime) As DateTime
            Try
                If Not String.IsNullOrWhiteSpace(filePath) AndAlso File.Exists(filePath) Then
                    Return File.GetLastWriteTime(filePath)
                End If
            Catch
            End Try

            Return fallback
        End Function

        Private Shared Function LocalizeMediaStatus(statusText As String) As String
            Select Case statusText
                Case "Success"
                    Return T("Shell.Media.Success", "Success")
                Case "Failed"
                    Return T("Shell.Media.Failed", "Failed")
                Case "Missing"
                    Return T("Shell.Media.Missing", "Missing")
                Case "Running"
                    Return T("Shell.Media.Running", "Running")
                Case "Pending"
                    Return T("Shell.Media.Pending", "Pending")
                Case Else
                    Return If(String.IsNullOrWhiteSpace(statusText), T("Project.Failure.None", "No issues"), statusText)
            End Select
        End Function

        Private Shared Function CreateBrush(hex As String) As Brush
            Dim brush = DirectCast(New BrushConverter().ConvertFromString(hex), Brush)
            brush.Freeze()
            Return brush
        End Function

        Private Sub Project_PropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            Dim project = TryCast(sender, ProjectSessionViewModel)
            If project Is Nothing Then Return

            RefreshFilteredCollections()
            NotifyDashboardChanged()

            If sender Is SelectedProject Then
                OnPropertyChanged(NameOf(WorkspaceTitle))
                OnPropertyChanged(NameOf(WorkspaceSubtitle))
                OnPropertyChanged(NameOf(CurrentPageDescription))
            End If

            If PersistedProjectProperties.Contains(e.PropertyName) Then
                PersistProject(project)
            End If
        End Sub

        Private Sub NotifyDashboardChanged()
            OnPropertyChanged(NameOf(ActiveProjectCount))
            OnPropertyChanged(NameOf(PlanningProjectCount))
            OnPropertyChanged(NameOf(RenderingProjectCount))
            OnPropertyChanged(NameOf(CompletedProjectCount))
            OnPropertyChanged(NameOf(FailedProjectCount))
            OnPropertyChanged(NameOf(AverageProgressValue))
            OnPropertyChanged(NameOf(AverageProgressText))
            OnPropertyChanged(NameOf(TotalMediaCount))
            OnPropertyChanged(NameOf(SourceMediaCount))
            OnPropertyChanged(NameOf(GeneratedMediaCount))
            OnPropertyChanged(NameOf(FailedMediaCount))
            OnPropertyChanged(NameOf(CurrentPageDescription))
            OnPropertyChanged(NameOf(StatusBarSummary))
        End Sub

        Private Sub NotifySystemSettingsModelBindingsChanged()
            OnPropertyChanged(NameOf(ModelSettings))
            OnPropertyChanged(NameOf(PlannerProviderEditors))
            OnPropertyChanged(NameOf(ImageProviderEditors))
            OnPropertyChanged(NameOf(DeepSeekPlannerModelOptionItems))
            OnPropertyChanged(NameOf(OpenAiPlannerModelOptionItems))
            OnPropertyChanged(NameOf(QwenPlannerModelOptionItems))
            OnPropertyChanged(NameOf(GeminiImageModelOptionItems))
            OnPropertyChanged(NameOf(OpenAiImageModelOptionItems))
            OnPropertyChanged(NameOf(QwenImageModelOptionItems))
        End Sub

        Private Sub NotifyPageStateChanged()
            OnPropertyChanged(NameOf(CurrentPageValue))
            OnPropertyChanged(NameOf(IsProjectHallPage))
            OnPropertyChanged(NameOf(IsWorkspacePage))
            OnPropertyChanged(NameOf(IsTaskQueuePage))
            OnPropertyChanged(NameOf(IsMediaLibraryPage))
            OnPropertyChanged(NameOf(IsSystemSettingsPage))
            OnPropertyChanged(NameOf(CurrentPageTitle))
            OnPropertyChanged(NameOf(CurrentPageDescription))
            OnPropertyChanged(NameOf(StatusBarSummary))
        End Sub

        Private Sub PersistProject(project As ProjectSessionViewModel)
            If _isRestoringState OrElse _isDisposed OrElse project Is Nothing Then Return
            _persistenceService.SaveProject(project.ToPersistedRecord())
        End Sub

        Private Sub SaveScenarioPreset(preset As ScenarioPreset)
            If _isRestoringState OrElse _isDisposed OrElse preset Is Nothing Then Return

            _persistenceService.SaveScenarioPreset(New PersistedScenarioPresetRecord With {
                .Id = preset.Id,
                .Name = preset.Name,
                .Description = preset.Description,
                .DesignPlanningTemplate = preset.DesignPlanningTemplate,
                .ImagePlanningTemplate = preset.ImagePlanningTemplate,
                .IsBuiltIn = preset.IsBuiltIn,
                .SortOrder = preset.SortOrder
            })
        End Sub

        Private Sub PersistStudioState()
            If _isRestoringState OrElse _isDisposed Then Return

            _persistenceService.SaveStudioState(New PersistedStudioState With {
                .SelectedProjectId = If(SelectedProject?.ProjectId, String.Empty),
                .CurrentPage = CInt(CurrentPage),
                .SearchText = SearchText,
                .ActiveFilter = CInt(ActiveFilter),
                .LanguageCode = LanguageCode
            })
        End Sub

        Public Sub NavigateToProjectHall()
            CurrentPage = StudioPage.ProjectHall
        End Sub

        Public Sub NavigateToWorkspace()
            If HasOpenWorkspaceProjects Then
                CurrentPage = StudioPage.ProjectWorkspace
            ElseIf SelectedProject IsNot Nothing Then
                OpenProject(SelectedProject)
            Else
                CurrentPage = StudioPage.ProjectHall
            End If
        End Sub

        Public Sub NavigateToTaskQueue()
            CurrentPage = StudioPage.TaskQueue
        End Sub

        Public Sub NavigateToMediaLibrary()
            CurrentPage = StudioPage.MediaLibrary
        End Sub

        Public Sub NavigateToSystemSettings()
            CurrentPage = StudioPage.SystemSettings
        End Sub

        Public Sub CreateNewProject()
            AddNewProject()
        End Sub

        Public Sub CreateScenarioPreset()
            Dim nextSortOrder = If(ScenarioPresets.Count = 0, 1, ScenarioPresets.Max(Function(item) item.SortOrder) + 1)
            Dim customScenarioCount = ScenarioPresets.Where(Function(item) Not item.IsBuiltIn).Count()
            Dim preset = New ScenarioPreset With {
                .Id = Guid.NewGuid().ToString("N"),
                .Name = $"自定义场景 {customScenarioCount + 1}",
                .Description = "可按当前业务内容自行补充描述。",
                .DesignPlanningTemplate = "请根据当前业务目标，先输出统一的设计规划，包括受众、视觉主题、配色、字体、摄影和版式策略。",
                .ImagePlanningTemplate = "请继续输出逐张图片规划，明确每张图的目的、视角、场景、卖点与文案承载方式。",
                .IsBuiltIn = False,
                .SortOrder = nextSortOrder
            }

            ScenarioPresets.Add(preset)
            SaveScenarioPreset(preset)
            SelectedScenarioPreset = preset
            CurrentPage = StudioPage.SystemSettings
        End Sub

        Public Sub DeleteSelectedScenarioPreset()
            Dim preset = SelectedScenarioPreset
            If preset Is Nothing Then Return
            If ScenarioPresets.Count <= 1 Then
                MessageBox.Show(T("Shell.Dialog.DeleteScenarioKeepOne", "At least one scenario template must remain."),
                                T("Shell.Dialog.DeleteScenario", "Delete Scenario"),
                                MessageBoxButton.OK,
                                MessageBoxImage.Information)
                Return
            End If

            Dim fallback = ScenarioPresets.FirstOrDefault(Function(item) Not String.Equals(item.Id, preset.Id, StringComparison.Ordinal))
            If fallback Is Nothing Then Return

            For Each project In Projects.Where(Function(item) String.Equals(item.SelectedScenarioPresetId, preset.Id, StringComparison.Ordinal))
                project.SelectedScenarioPresetId = fallback.Id
                PersistProject(project)
            Next

            ScenarioPresets.Remove(preset)
            _persistenceService.DeleteScenarioPreset(preset.Id)
            SelectedScenarioPreset = fallback
        End Sub

        Public Sub SaveSelectedScenarioPreset()
            Dim preset = SelectedScenarioPreset
            If preset Is Nothing Then
                Throw New InvalidOperationException("当前没有可保存的场景模板。")
            End If

            SaveScenarioPreset(preset)

            For Each project In Projects
                project.NotifyScenarioPresetsChanged()
            Next

            RefreshFilteredCollections()
        End Sub

        Public Sub SaveModelSettings()
            If ModelSettings Is Nothing Then
                Throw New InvalidOperationException("当前没有可保存的模型配置。")
            End If

            VertexAiService.SaveModelProviderEditors(PlannerProviderEditors, ImageProviderEditors)
            ModelSettings = VertexAiService.LoadModelConfiguration()
            _vertexService.ReloadConfiguration()

            For Each project In Projects
                project.RefreshProviderConfiguration()
                PersistProject(project)
            Next
        End Sub

        Public Sub SetAllFilter()
            ActiveFilter = ProjectFilter.All
        End Sub

        Public Sub SetPlanningFilter()
            ActiveFilter = ProjectFilter.Planning
        End Sub

        Public Sub SetRenderingFilter()
            ActiveFilter = ProjectFilter.Rendering
        End Sub

        Public Sub SetReadyFilter()
            ActiveFilter = ProjectFilter.Ready
        End Sub

        Public Sub SetFailedFilter()
            ActiveFilter = ProjectFilter.Failed
        End Sub

        Public Sub SetCompletedFilter()
            ActiveFilter = ProjectFilter.Completed
        End Sub

        Public Sub SetAllMediaFilter()
            ActiveMediaFilter = MediaLibraryFilter.All
        End Sub

        Public Sub SetSourceMediaFilter()
            ActiveMediaFilter = MediaLibraryFilter.Source
        End Sub

        Public Sub SetGeneratedMediaFilter()
            ActiveMediaFilter = MediaLibraryFilter.Generated
        End Sub

        Public Sub SetFailedMediaFilter()
            ActiveMediaFilter = MediaLibraryFilter.Failed
        End Sub

        Public Sub OpenProject(project As ProjectSessionViewModel)
            If project Is Nothing Then Return
            EnsureWorkspaceTab(project)
            SelectedProject = project
            CurrentPage = StudioPage.ProjectWorkspace
        End Sub

        Public Sub OpenMediaLibraryProject(item As MediaLibraryItem)
            If item?.OwnerProject Is Nothing Then Return
            OpenProject(item.OwnerProject)
        End Sub

        Public Sub OpenSelectedMediaLibraryProject()
            OpenMediaLibraryProject(SelectedMediaLibraryItem)
        End Sub

        Public Sub CloseWorkspaceTab(project As ProjectSessionViewModel)
            If project Is Nothing OrElse Not OpenWorkspaceProjects.Contains(project) Then Return

            Dim tabIndex = OpenWorkspaceProjects.IndexOf(project)
            Dim wasSelected = ReferenceEquals(SelectedProject, project)
            OpenWorkspaceProjects.Remove(project)

            If OpenWorkspaceProjects.Count = 0 Then
                If CurrentPage = StudioPage.ProjectWorkspace Then
                    CurrentPage = StudioPage.ProjectHall
                End If
            ElseIf wasSelected Then
                Dim nextIndex = Math.Max(0, Math.Min(tabIndex, OpenWorkspaceProjects.Count - 1))
                SelectedProject = OpenWorkspaceProjects(nextIndex)
                CurrentPage = StudioPage.ProjectWorkspace
            End If

            PersistStudioState()
        End Sub

        Public Sub CloseProject(project As ProjectSessionViewModel)
            If project Is Nothing OrElse Not Projects.Contains(project) Then Return

            If project.IsBusy Then
                Dim result = MessageBox.Show(F("Shell.Dialog.CloseProjectBusy", "{0} is still running. Closing it will stop the current task. Continue?", project.ProjectName),
                                             T("Shell.Dialog.CloseProject", "Close Project"),
                                             MessageBoxButton.YesNo,
                                             MessageBoxImage.Question)
                If result <> MessageBoxResult.Yes Then Return
                project.StopCurrentOperation("项目已取消。")
            End If

            Dim tabIndex = OpenWorkspaceProjects.IndexOf(project)
            Dim wasSelected = ReferenceEquals(SelectedProject, project)
            project.IsArchived = True
            If OpenWorkspaceProjects.Contains(project) Then
                OpenWorkspaceProjects.Remove(project)
            End If
            PersistProject(project)

            If wasSelected Then
                If OpenWorkspaceProjects.Count > 0 Then
                    Dim nextTabIndex = Math.Max(0, Math.Min(tabIndex, OpenWorkspaceProjects.Count - 1))
                    SelectedProject = OpenWorkspaceProjects(nextTabIndex)
                    If CurrentPage = StudioPage.ProjectWorkspace Then
                        CurrentPage = StudioPage.ProjectWorkspace
                    End If
                Else
                    SelectedProject = Projects.FirstOrDefault(Function(item) Not item.IsArchived)
                    If CurrentPage = StudioPage.ProjectWorkspace Then
                        CurrentPage = StudioPage.ProjectHall
                    End If
                End If
            End If

            RebuildMediaLibrary()
            RefreshFilteredCollections()
            NotifyDashboardChanged()
            PersistStudioState()
        End Sub

        Public Async Function ExecuteQueueActionAsync(project As ProjectSessionViewModel) As Task
            If project Is Nothing Then Return

            Try
                If project.IsPlanning Then
                    Await project.ToggleGeneratePlanAsync()
                    PersistProject(project)
                    Return
                End If

                If project.IsGeneratingImages Then
                    Await project.ToggleGenerateImagesAsync()
                    PersistProject(project)
                    Return
                End If

                If project.HasRetryableFailures Then
                    Await RetryFailedImagesForProjectAsync(project)
                    Return
                End If

                If CanStartProjectTask(project) Then
                    Await StartProjectTaskWithLimitAsync(project)
                    Return
                End If

                OpenProject(project)
            Catch ex As Exception
                MessageBox.Show(ex.Message, T("Shell.Dialog.TaskExecutionFailed", "Task execution failed"), MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Function

        Public Async Function StartSelectedQueueProjectsAsync() As Task
            Dim selectedProjects = Projects.
                Where(Function(project) project.IsQueueSelected AndAlso CanStartProjectTask(project)).
                ToList()

            If selectedProjects.Count = 0 Then Return

            Await Task.WhenAll(selectedProjects.Select(Function(project) StartProjectTaskWithLimitAsync(project)))
        End Function

        Private Function CanStartProjectTask(project As ProjectSessionViewModel) As Boolean
            If project Is Nothing Then Return False
            If project.IsPlanning OrElse project.IsGeneratingImages Then Return False
            If project.HasRetryableFailures Then Return False

            If project.UploadedImages.Count > 0 AndAlso project.HasGeneratedPlan Then
                Return True
            End If

            If project.UploadedImages.Count > 0 Then
                Return True
            End If

            Return False
        End Function

        Private Async Function StartProjectTaskAsync(project As ProjectSessionViewModel) As Task
            If project Is Nothing Then Return

            If project.UploadedImages.Count > 0 AndAlso project.HasGeneratedPlan Then
                Await project.ToggleGenerateImagesAsync()
                PersistProject(project)
                Return
            End If

            If project.UploadedImages.Count > 0 Then
                Await project.ToggleGeneratePlanAsync()
                PersistProject(project)
            End If
        End Function

        Private Async Function StartProjectTaskWithLimitAsync(project As ProjectSessionViewModel) As Task
            If project Is Nothing Then Return

            Await _taskExecutionSemaphore.WaitAsync()
            Try
                Await StartProjectTaskAsync(project)
            Finally
                _taskExecutionSemaphore.Release()
            End Try
        End Function

        Public Async Function RetryFailedImagesForProjectAsync(project As ProjectSessionViewModel) As Task
            If project Is Nothing Then Return

            Try
                Dim retryItems = project.GeneratedImages.
                    Where(Function(item) String.Equals(item.StatusText, "Failed", StringComparison.Ordinal) OrElse
                                         String.Equals(item.StatusText, "Missing", StringComparison.Ordinal)).
                    OrderBy(Function(item) item.Sequence).
                    ToList()

                If retryItems.Count = 0 Then
                    OpenProject(project)
                    Return
                End If

                SelectedProject = project
                For Each item In retryItems
                    Await project.RegenerateImageAsync(item)
                Next
                PersistProject(project)
            Catch ex As Exception
                MessageBox.Show(ex.Message, T("Shell.Dialog.RetryFailedImages", "Retry failed images failed"), MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Function

        Public Async Function RetryFailedImagesForFilteredProjectsAsync() As Task
            Dim retryProjects = FilteredTaskProjects.
                Cast(Of Object)().
                OfType(Of ProjectSessionViewModel)().
                Where(Function(project) project.HasRetryableFailures).
                ToList()

            If retryProjects.Count = 0 Then Return

            For Each project In retryProjects
                Await RetryFailedImagesForProjectAsync(project)
            Next
        End Function

        Public Sub ImportImagesForSelectedProject()
            If SelectedProject Is Nothing Then Return

            Dim dialog = New OpenFileDialog With {
                .Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp;*.webp",
                .Multiselect = True,
                .Title = "导入产品图"
            }

            If dialog.ShowDialog() <> True Then Return
            SelectedProject.SetUploadedImages(dialog.FileNames)
            PersistProject(SelectedProject)
        End Sub

        Public Sub RemoveUploadedImageFromSelectedProject(item As ProductImageItem)
            If SelectedProject Is Nothing Then Return
            SelectedProject.RemoveUploadedImage(item)
            PersistProject(SelectedProject)
        End Sub

        Public Async Function GeneratePlanForSelectedProjectAsync() As Task
            If SelectedProject Is Nothing Then Return

            Try
                Await SelectedProject.ToggleGeneratePlanAsync()
                PersistProject(SelectedProject)
            Catch ex As Exception
                MessageBox.Show(ex.Message, T("Shell.Dialog.PlanFailed", "Planning failed"), MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Function

        Public Async Function GenerateImagesForSelectedProjectAsync() As Task
            If SelectedProject Is Nothing Then Return

            Try
                Await SelectedProject.ToggleGenerateImagesAsync()
                PersistProject(SelectedProject)
            Catch ex As Exception
                MessageBox.Show(ex.Message, T("Shell.Dialog.RenderFailed", "Image generation failed"), MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Function

        Public Async Function RegenerateImageForSelectedProjectAsync(item As GeneratedImageItem) As Task
            If SelectedProject Is Nothing OrElse item Is Nothing Then Return

            Try
                Await SelectedProject.RegenerateImageAsync(item)
                PersistProject(SelectedProject)
            Catch ex As Exception
                MessageBox.Show(ex.Message, T("Shell.Dialog.RegenerateFailed", "Regeneration failed"), MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Function

        Public Sub DownloadImageItem(item As GeneratedImageItem)
            If item Is Nothing OrElse String.IsNullOrWhiteSpace(item.SavedPath) OrElse Not File.Exists(item.SavedPath) Then Return

            Dim dialog = New SaveFileDialog With {
                .Title = "下载图片",
                .FileName = Path.GetFileName(item.SavedPath),
                .Filter = "PNG 图片|*.png|JPEG 图片|*.jpg;*.jpeg|所有文件|*.*"
            }

            If dialog.ShowDialog() <> True Then Return
            File.Copy(item.SavedPath, dialog.FileName, overwrite:=True)
        End Sub

        Public Sub DownloadMediaLibraryItem(item As MediaLibraryItem)
            If item Is Nothing OrElse String.IsNullOrWhiteSpace(item.FilePath) OrElse Not File.Exists(item.FilePath) Then Return

            Dim dialog = New SaveFileDialog With {
                .Title = "下载图片",
                .FileName = Path.GetFileName(item.FilePath),
                .Filter = "图片文件|*.png;*.jpg;*.jpeg;*.bmp;*.webp|所有文件|*.*"
            }

            If dialog.ShowDialog() <> True Then Return
            File.Copy(item.FilePath, dialog.FileName, overwrite:=True)
        End Sub

        Public Sub EditImageItem(item As GeneratedImageItem)
            If item Is Nothing OrElse String.IsNullOrWhiteSpace(item.SavedPath) OrElse Not File.Exists(item.SavedPath) Then Return
            Process.Start(New ProcessStartInfo With {.FileName = item.SavedPath, .UseShellExecute = True})
        End Sub

        Public Sub OpenMediaLibraryFile(item As MediaLibraryItem)
            If item Is Nothing OrElse String.IsNullOrWhiteSpace(item.FilePath) OrElse Not File.Exists(item.FilePath) Then Return
            Process.Start(New ProcessStartInfo With {.FileName = item.FilePath, .UseShellExecute = True})
        End Sub

        Public Sub OpenSelectedMediaLibraryFile()
            OpenMediaLibraryFile(SelectedMediaLibraryItem)
        End Sub

        Public Sub OpenSelectedProjectOutputDirectory()
            If SelectedProject Is Nothing Then
                MessageBox.Show(T("Shell.Dialog.NoSelectedProject", "No project is currently selected."),
                                T("Shell.Dialog.OpenOutput", "Open Output Folder"),
                                MessageBoxButton.OK,
                                MessageBoxImage.Information)
                Return
            End If

            Dim outputDirectory = ResolveOutputDirectory(SelectedProject)
            If String.IsNullOrWhiteSpace(outputDirectory) Then
                MessageBox.Show(T("Shell.Dialog.NoOutputFolder", "This project does not have an output folder yet. Generate images first."),
                                T("Shell.Dialog.OpenOutput", "Open Output Folder"),
                                MessageBoxButton.OK,
                                MessageBoxImage.Information)
                Return
            End If

            Try
                Process.Start(New ProcessStartInfo With {
                    .FileName = outputDirectory,
                    .UseShellExecute = True,
                    .Verb = "open",
                    .WorkingDirectory = outputDirectory
                })
            Catch
                Try
                    Process.Start(New ProcessStartInfo With {
                        .FileName = "explorer.exe",
                        .Arguments = $"""{outputDirectory}""",
                        .UseShellExecute = True,
                        .WorkingDirectory = outputDirectory
                    })
                Catch ex As Exception
                    MessageBox.Show(F("Shell.Dialog.OpenOutputFailed", "Failed to open output folder: {0}", ex.Message),
                                    T("Shell.Dialog.OpenOutput", "Open Output Folder"),
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Error)
                End Try
            End Try
        End Sub

        Private Shared Function ResolveOutputDirectory(project As ProjectSessionViewModel) As String
            If project Is Nothing Then
                Return String.Empty
            End If

            Dim preferredDirectory = If(project.LastOutputDirectory, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(preferredDirectory) Then
                Try
                    preferredDirectory = Path.GetFullPath(preferredDirectory)
                Catch
                    preferredDirectory = preferredDirectory.Trim()
                End Try

                If Directory.Exists(preferredDirectory) Then
                    If Not String.Equals(project.LastOutputDirectory, preferredDirectory, StringComparison.OrdinalIgnoreCase) Then
                        project.LastOutputDirectory = preferredDirectory
                    End If

                    Return preferredDirectory
                End If
            End If

            Dim derivedDirectory = project.GeneratedImages.
                Select(Function(item) If(item?.SavedPath, String.Empty)).
                Where(Function(savedPath) Not String.IsNullOrWhiteSpace(savedPath) AndAlso File.Exists(savedPath)).
                Select(Function(savedPath) IO.Path.GetDirectoryName(savedPath)).
                FirstOrDefault(Function(directoryPath) Not String.IsNullOrWhiteSpace(directoryPath) AndAlso Directory.Exists(directoryPath))

            If Not String.IsNullOrWhiteSpace(derivedDirectory) Then
                project.LastOutputDirectory = derivedDirectory
                Return derivedDirectory
            End If

            Return String.Empty
        End Function

        Public Sub PreviewGeneratedImage(item As GeneratedImageItem, owner As Window)
            If item Is Nothing OrElse Not item.HasPreview Then Return

            Dim previewWindow As New Window With {
                .Title = item.Title,
                .Width = 1180,
                .Height = 820,
                .MinWidth = 780,
                .MinHeight = 560,
                .WindowStartupLocation = WindowStartupLocation.CenterOwner,
                .Owner = owner,
                .WindowStyle = WindowStyle.None,
                .ResizeMode = ResizeMode.NoResize,
                .AllowsTransparency = True,
                .Background = Brushes.Transparent,
                .ShowInTaskbar = False
            }

            Dim rootGrid As New Grid With {
                .Background = New SolidColorBrush(Color.FromArgb(&HE6, &H9, &HE, &H1A))
            }

            AddHandler previewWindow.KeyDown,
                Sub(sender2, e2)
                    If e2.Key = Key.Escape Then previewWindow.Close()
                End Sub

            Dim dialogBorder As New Border With {
                .Width = 1040,
                .MaxWidth = 1040,
                .MaxHeight = 720,
                .Padding = New Thickness(18),
                .Background = New SolidColorBrush(Color.FromRgb(&H11, &H18, &H27)),
                .BorderBrush = New SolidColorBrush(Color.FromRgb(&H2E, &H3A, &H4A)),
                .BorderThickness = New Thickness(1),
                .CornerRadius = New CornerRadius(20),
                .HorizontalAlignment = HorizontalAlignment.Center,
                .VerticalAlignment = VerticalAlignment.Center
            }

            Dim dialogGrid As New Grid()
            dialogGrid.RowDefinitions.Add(New RowDefinition With {.Height = GridLength.Auto})
            dialogGrid.RowDefinitions.Add(New RowDefinition With {.Height = New GridLength(1, GridUnitType.Star)})

            Dim titleGrid As New Grid With {.Margin = New Thickness(4, 0, 4, 12)}
            titleGrid.ColumnDefinitions.Add(New ColumnDefinition With {.Width = New GridLength(1, GridUnitType.Star)})
            titleGrid.ColumnDefinitions.Add(New ColumnDefinition With {.Width = GridLength.Auto})

            Dim titlePanel As New StackPanel()
            titlePanel.Children.Add(New TextBlock With {
                .Text = item.Title,
                .Foreground = Brushes.White,
                .FontSize = 22,
                .FontFamily = New FontFamily("Microsoft YaHei UI Semibold")
            })
            titlePanel.Children.Add(New TextBlock With {
                .Text = "按 Esc 或右上角关闭",
                .Foreground = New SolidColorBrush(Color.FromRgb(&H94, &HA3, &HB8)),
                .FontSize = 12,
                .Margin = New Thickness(0, 6, 0, 0)
            })

            Dim closeButton As New Button With {
                .Width = 34,
                .Height = 34,
                .Content = "X",
                .FontSize = 14,
                .FontFamily = New FontFamily("Segoe UI Semibold"),
                .Background = New SolidColorBrush(Color.FromRgb(&H17, &H22, &H33)),
                .Foreground = Brushes.White,
                .BorderBrush = New SolidColorBrush(Color.FromRgb(&H3B, &H4A, &H5E)),
                .BorderThickness = New Thickness(1),
                .Padding = New Thickness(0),
                .Cursor = Cursors.Hand,
                .HorizontalAlignment = HorizontalAlignment.Right,
                .VerticalAlignment = VerticalAlignment.Top
            }
            AddHandler closeButton.Click, Sub() previewWindow.Close()

            Grid.SetColumn(titlePanel, 0)
            Grid.SetColumn(closeButton, 1)
            titleGrid.Children.Add(titlePanel)
            titleGrid.Children.Add(closeButton)

            Dim imageBorder As New Border With {
                .Background = New SolidColorBrush(Color.FromRgb(&HF8, &HFA, &HFD)),
                .CornerRadius = New CornerRadius(16),
                .Padding = New Thickness(12)
            }
            imageBorder.Child = New Image With {.Source = item.Preview, .Stretch = Stretch.Uniform}

            Grid.SetRow(titleGrid, 0)
            Grid.SetRow(imageBorder, 1)
            dialogGrid.Children.Add(titleGrid)
            dialogGrid.Children.Add(imageBorder)
            dialogBorder.Child = dialogGrid
            rootGrid.Children.Add(dialogBorder)
            previewWindow.Content = rootGrid
            previewWindow.ShowDialog()
        End Sub

        Public Sub PreviewMediaLibraryItem(item As MediaLibraryItem, owner As Window)
            If item Is Nothing OrElse Not item.HasPreview Then Return

            Dim previewWindow As New Window With {
                .Title = item.Title,
                .Width = 1180,
                .Height = 820,
                .MinWidth = 780,
                .MinHeight = 560,
                .WindowStartupLocation = WindowStartupLocation.CenterOwner,
                .Owner = owner,
                .WindowStyle = WindowStyle.None,
                .ResizeMode = ResizeMode.NoResize,
                .AllowsTransparency = True,
                .Background = Brushes.Transparent,
                .ShowInTaskbar = False
            }

            Dim rootGrid As New Grid With {
                .Background = New SolidColorBrush(Color.FromArgb(&HE6, &H9, &HE, &H1A))
            }

            AddHandler previewWindow.KeyDown,
                Sub(sender2, e2)
                    If e2.Key = Key.Escape Then previewWindow.Close()
                End Sub

            Dim dialogBorder As New Border With {
                .Width = 1040,
                .MaxWidth = 1040,
                .MaxHeight = 720,
                .Padding = New Thickness(18),
                .Background = New SolidColorBrush(Color.FromRgb(&H11, &H18, &H27)),
                .BorderBrush = New SolidColorBrush(Color.FromRgb(&H2E, &H3A, &H4A)),
                .BorderThickness = New Thickness(1),
                .CornerRadius = New CornerRadius(20),
                .HorizontalAlignment = HorizontalAlignment.Center,
                .VerticalAlignment = VerticalAlignment.Center
            }

            Dim dialogGrid As New Grid()
            dialogGrid.RowDefinitions.Add(New RowDefinition With {.Height = GridLength.Auto})
            dialogGrid.RowDefinitions.Add(New RowDefinition With {.Height = New GridLength(1, GridUnitType.Star)})

            Dim titleGrid As New Grid With {.Margin = New Thickness(4, 0, 4, 12)}
            titleGrid.ColumnDefinitions.Add(New ColumnDefinition With {.Width = New GridLength(1, GridUnitType.Star)})
            titleGrid.ColumnDefinitions.Add(New ColumnDefinition With {.Width = GridLength.Auto})

            Dim titlePanel As New StackPanel()
            titlePanel.Children.Add(New TextBlock With {
                .Text = item.Title,
                .Foreground = Brushes.White,
                .FontSize = 22,
                .FontFamily = New FontFamily("Microsoft YaHei UI Semibold")
            })
            titlePanel.Children.Add(New TextBlock With {
                .Text = $"{item.ProjectName} · {item.KindText}",
                .Foreground = New SolidColorBrush(Color.FromRgb(&H94, &HA3, &HB8)),
                .FontSize = 12,
                .Margin = New Thickness(0, 6, 0, 0)
            })

            Dim closeButton As New Button With {
                .Width = 34,
                .Height = 34,
                .Content = "X",
                .FontSize = 14,
                .FontFamily = New FontFamily("Segoe UI Semibold"),
                .Background = New SolidColorBrush(Color.FromRgb(&H17, &H22, &H33)),
                .Foreground = Brushes.White,
                .BorderBrush = New SolidColorBrush(Color.FromRgb(&H3B, &H4A, &H5E)),
                .BorderThickness = New Thickness(1),
                .Padding = New Thickness(0),
                .Cursor = Cursors.Hand,
                .HorizontalAlignment = HorizontalAlignment.Right,
                .VerticalAlignment = VerticalAlignment.Top
            }
            AddHandler closeButton.Click, Sub() previewWindow.Close()

            Grid.SetColumn(titlePanel, 0)
            Grid.SetColumn(closeButton, 1)
            titleGrid.Children.Add(titlePanel)
            titleGrid.Children.Add(closeButton)

            Dim imageBorder As New Border With {
                .Background = New SolidColorBrush(Color.FromRgb(&HF8, &HFA, &HFD)),
                .CornerRadius = New CornerRadius(16),
                .Padding = New Thickness(12)
            }
            imageBorder.Child = New Image With {.Source = item.Preview, .Stretch = Stretch.Uniform}

            Grid.SetRow(titleGrid, 0)
            Grid.SetRow(imageBorder, 1)
            dialogGrid.Children.Add(titleGrid)
            dialogGrid.Children.Add(imageBorder)
            dialogBorder.Child = dialogGrid
            rootGrid.Children.Add(dialogBorder)
            previewWindow.Content = rootGrid
            previewWindow.ShowDialog()
        End Sub

        Public Sub PreviewSelectedMediaLibraryItem(owner As Window)
            PreviewMediaLibraryItem(SelectedMediaLibraryItem, owner)
        End Sub

        Public Sub OpenProjectFromQueue(project As ProjectSessionViewModel)
            If project Is Nothing Then Return
            OpenProject(project)
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _isDisposed Then Return
            _isDisposed = True

            PersistStudioState()
            For Each preset In ScenarioPresets.ToList()
                SaveScenarioPreset(preset)
                DetachScenarioPreset(preset)
            Next
            For Each project In Projects.ToList()
                PersistProject(project)
                DetachProject(project)
                project.Dispose()
            Next

            _taskExecutionSemaphore.Dispose()
        End Sub
    End Class
End Namespace


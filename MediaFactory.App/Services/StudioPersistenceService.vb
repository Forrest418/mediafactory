Imports System
Imports System.Globalization
Imports System.IO
Imports System.Text.Json
Imports System.Text.Json.Serialization.Metadata
Imports Microsoft.Data.Sqlite
Imports MediaFactory.Models

Namespace Services
    Public Class StudioPersistenceService
        Private Const BuiltInScenarioTemplateVersion As Integer = 4

        Private Shared ReadOnly JsonOptions As New JsonSerializerOptions With {
            .WriteIndented = False,
            .TypeInfoResolver = New DefaultJsonTypeInfoResolver()
        }

        Private ReadOnly _databasePath As String

        Public Sub New()
            Dim dataDirectory = Path.Combine(AppContext.BaseDirectory, "Data")
            Directory.CreateDirectory(dataDirectory)
            _databasePath = Path.Combine(dataDirectory, "studio.db")
            EnsureSchema()
        End Sub

        Public ReadOnly Property DatabasePath As String
            Get
                Return _databasePath
            End Get
        End Property

        Public Function LoadProjects() As List(Of PersistedProjectRecord)
            Dim results As New List(Of PersistedProjectRecord)()

            Using connection = CreateConnection()
                connection.Open()

                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT project_id, project_name, is_archived, project_summary, selected_scenario_preset_id, selected_planner_provider_key, requirement_text, target_language, selected_planner_model, selected_image_provider_key, selected_image_model, selected_aspect_ratio, output_resolution, requested_count, derivative_image_count, auto_render_after_plan, design_plan_markdown, image_plan_markdown, has_generated_plan, last_output_directory, last_activated_at, uploaded_images_json, generated_images_json " &
                        "FROM project_sessions ORDER BY last_activated_at DESC, project_name ASC;"

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            results.Add(New PersistedProjectRecord With {
                                .ProjectId = reader.GetString(0),
                                .ProjectName = reader.GetString(1),
                                .IsArchived = reader.GetInt32(2) = 1,
                                .ProjectSummary = reader.GetString(3),
                                .SelectedScenarioPresetId = reader.GetString(4),
                                .SelectedPlannerProviderKey = reader.GetString(5),
                                .RequirementText = reader.GetString(6),
                                .TargetLanguage = reader.GetString(7),
                                .SelectedPlannerModel = reader.GetString(8),
                                .SelectedImageProviderKey = reader.GetString(9),
                                .SelectedImageModel = reader.GetString(10),
                                .SelectedAspectRatio = reader.GetString(11),
                                .OutputResolution = reader.GetString(12),
                                .RequestedCount = reader.GetInt32(13),
                                .DerivativeImageCount = reader.GetInt32(14),
                                .AutoRenderAfterPlan = reader.GetInt32(15) = 1,
                                .DesignPlanMarkdown = reader.GetString(16),
                                .ImagePlanMarkdown = reader.GetString(17),
                                .HasGeneratedPlan = reader.GetInt32(18) = 1,
                                .LastOutputDirectory = reader.GetString(19),
                                .LastActivatedAt = ParseRoundtripDateTime(reader.GetString(20)),
                                .UploadedImagePaths = DeserializeList(Of String)(reader.GetString(21)),
                                .GeneratedImages = DeserializeList(Of PersistedGeneratedImageRecord)(reader.GetString(22))
                            })
                        End While
                    End Using
                End Using
            End Using

            Return results
        End Function

        Public Function LoadScenarioPresets() As List(Of PersistedScenarioPresetRecord)
            Dim results As New List(Of PersistedScenarioPresetRecord)()

            Using connection = CreateConnection()
                connection.Open()

                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT preset_id, name, description, planning_instruction, image_instruction, is_builtin, sort_order " &
                        "FROM scenario_presets ORDER BY sort_order ASC, name ASC;"

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            results.Add(New PersistedScenarioPresetRecord With {
                                .Id = reader.GetString(0),
                                .Name = reader.GetString(1),
                                .Description = reader.GetString(2),
                                .DesignPlanningTemplate = reader.GetString(3),
                                .ImagePlanningTemplate = reader.GetString(4),
                                .IsBuiltIn = reader.GetInt32(5) = 1,
                                .SortOrder = reader.GetInt32(6)
                            })
                        End While
                    End Using
                End Using
            End Using

            Return results
        End Function

        Public Function LoadStudioState() As PersistedStudioState
            Using connection = CreateConnection()
                connection.Open()

                Using command = connection.CreateCommand()
                    command.CommandText = "SELECT selected_project_id, current_page, search_text, active_filter, language_code FROM studio_state WHERE id = 1;"
                    Using reader = command.ExecuteReader()
                        If reader.Read() Then
                            Return New PersistedStudioState With {
                                .SelectedProjectId = reader.GetString(0),
                                .CurrentPage = reader.GetInt32(1),
                                .SearchText = reader.GetString(2),
                                .ActiveFilter = reader.GetInt32(3),
                                .LanguageCode = If(reader.IsDBNull(4), "en", reader.GetString(4))
                            }
                        End If
                    End Using
                End Using
            End Using

            Return New PersistedStudioState()
        End Function

        Public Sub SaveProject(record As PersistedProjectRecord)
            Using connection = CreateConnection()
                connection.Open()

                Using command = connection.CreateCommand()
                    command.CommandText =
                        "INSERT INTO project_sessions (project_id, project_name, is_archived, project_summary, selected_scenario_preset_id, selected_planner_provider_key, requirement_text, target_language, selected_planner_model, selected_image_provider_key, selected_image_model, selected_aspect_ratio, output_resolution, requested_count, derivative_image_count, auto_render_after_plan, design_plan_markdown, image_plan_markdown, has_generated_plan, last_output_directory, last_activated_at, uploaded_images_json, generated_images_json) " &
                        "VALUES ($projectId, $projectName, $isArchived, $projectSummary, $selectedScenarioPresetId, $selectedPlannerProviderKey, $requirementText, $targetLanguage, $selectedPlannerModel, $selectedImageProviderKey, $selectedImageModel, $selectedAspectRatio, $outputResolution, $requestedCount, $derivativeImageCount, $autoRenderAfterPlan, $designPlanMarkdown, $imagePlanMarkdown, $hasGeneratedPlan, $lastOutputDirectory, $lastActivatedAt, $uploadedImagesJson, $generatedImagesJson) " &
                        "ON CONFLICT(project_id) DO UPDATE SET " &
                        "project_name = excluded.project_name, " &
                        "is_archived = excluded.is_archived, " &
                        "project_summary = excluded.project_summary, " &
                        "selected_scenario_preset_id = excluded.selected_scenario_preset_id, " &
                        "selected_planner_provider_key = excluded.selected_planner_provider_key, " &
                        "requirement_text = excluded.requirement_text, " &
                        "target_language = excluded.target_language, " &
                        "selected_planner_model = excluded.selected_planner_model, " &
                        "selected_image_provider_key = excluded.selected_image_provider_key, " &
                        "selected_image_model = excluded.selected_image_model, " &
                        "selected_aspect_ratio = excluded.selected_aspect_ratio, " &
                        "output_resolution = excluded.output_resolution, " &
                        "requested_count = excluded.requested_count, " &
                        "derivative_image_count = excluded.derivative_image_count, " &
                        "auto_render_after_plan = excluded.auto_render_after_plan, " &
                        "design_plan_markdown = excluded.design_plan_markdown, " &
                        "image_plan_markdown = excluded.image_plan_markdown, " &
                        "has_generated_plan = excluded.has_generated_plan, " &
                        "last_output_directory = excluded.last_output_directory, " &
                        "last_activated_at = excluded.last_activated_at, " &
                        "uploaded_images_json = excluded.uploaded_images_json, " &
                        "generated_images_json = excluded.generated_images_json;"

                    command.Parameters.AddWithValue("$projectId", record.ProjectId)
                    command.Parameters.AddWithValue("$projectName", record.ProjectName)
                    command.Parameters.AddWithValue("$isArchived", If(record.IsArchived, 1, 0))
                    command.Parameters.AddWithValue("$projectSummary", record.ProjectSummary)
                    command.Parameters.AddWithValue("$selectedScenarioPresetId", record.SelectedScenarioPresetId)
                    command.Parameters.AddWithValue("$selectedPlannerProviderKey", record.SelectedPlannerProviderKey)
                    command.Parameters.AddWithValue("$requirementText", record.RequirementText)
                    command.Parameters.AddWithValue("$targetLanguage", record.TargetLanguage)
                    command.Parameters.AddWithValue("$selectedPlannerModel", record.SelectedPlannerModel)
                    command.Parameters.AddWithValue("$selectedImageProviderKey", record.SelectedImageProviderKey)
                    command.Parameters.AddWithValue("$selectedImageModel", record.SelectedImageModel)
                    command.Parameters.AddWithValue("$selectedAspectRatio", record.SelectedAspectRatio)
                    command.Parameters.AddWithValue("$outputResolution", record.OutputResolution)
                    command.Parameters.AddWithValue("$requestedCount", record.RequestedCount)
                    command.Parameters.AddWithValue("$derivativeImageCount", record.DerivativeImageCount)
                    command.Parameters.AddWithValue("$autoRenderAfterPlan", If(record.AutoRenderAfterPlan, 1, 0))
                    command.Parameters.AddWithValue("$designPlanMarkdown", record.DesignPlanMarkdown)
                    command.Parameters.AddWithValue("$imagePlanMarkdown", record.ImagePlanMarkdown)
                    command.Parameters.AddWithValue("$hasGeneratedPlan", If(record.HasGeneratedPlan, 1, 0))
                    command.Parameters.AddWithValue("$lastOutputDirectory", record.LastOutputDirectory)
                    command.Parameters.AddWithValue("$lastActivatedAt", record.LastActivatedAt.ToString("O"))
                    command.Parameters.AddWithValue("$uploadedImagesJson", JsonSerializer.Serialize(record.UploadedImagePaths, JsonOptions))
                    command.Parameters.AddWithValue("$generatedImagesJson", JsonSerializer.Serialize(record.GeneratedImages, JsonOptions))
                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub SaveScenarioPreset(record As PersistedScenarioPresetRecord)
            Using connection = CreateConnection()
                connection.Open()

                Using command = connection.CreateCommand()
                    command.CommandText =
                        "INSERT INTO scenario_presets (preset_id, name, description, planning_instruction, image_instruction, is_builtin, sort_order) " &
                        "VALUES ($presetId, $name, $description, $planningInstruction, $imageInstruction, $isBuiltIn, $sortOrder) " &
                        "ON CONFLICT(preset_id) DO UPDATE SET " &
                        "name = excluded.name, " &
                        "description = excluded.description, " &
                        "planning_instruction = excluded.planning_instruction, " &
                        "image_instruction = excluded.image_instruction, " &
                        "is_builtin = excluded.is_builtin, " &
                        "sort_order = excluded.sort_order;"

                    command.Parameters.AddWithValue("$presetId", record.Id)
                    command.Parameters.AddWithValue("$name", record.Name)
                    command.Parameters.AddWithValue("$description", record.Description)
                    command.Parameters.AddWithValue("$planningInstruction", record.DesignPlanningTemplate)
                    command.Parameters.AddWithValue("$imageInstruction", record.ImagePlanningTemplate)
                    command.Parameters.AddWithValue("$isBuiltIn", If(record.IsBuiltIn, 1, 0))
                    command.Parameters.AddWithValue("$sortOrder", record.SortOrder)
                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub DeleteScenarioPreset(presetId As String)
            Using connection = CreateConnection()
                connection.Open()

                Using command = connection.CreateCommand()
                    command.CommandText = "DELETE FROM scenario_presets WHERE preset_id = $presetId;"
                    command.Parameters.AddWithValue("$presetId", presetId)
                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub DeleteProject(projectId As String)
            Using connection = CreateConnection()
                connection.Open()

                Using command = connection.CreateCommand()
                    command.CommandText = "DELETE FROM project_sessions WHERE project_id = $projectId;"
                    command.Parameters.AddWithValue("$projectId", projectId)
                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub SaveStudioState(state As PersistedStudioState)
            Using connection = CreateConnection()
                connection.Open()

                Using command = connection.CreateCommand()
                    command.CommandText =
                        "INSERT INTO studio_state (id, selected_project_id, current_page, search_text, active_filter, language_code) " &
                        "VALUES (1, $selectedProjectId, $currentPage, $searchText, $activeFilter, $languageCode) " &
                        "ON CONFLICT(id) DO UPDATE SET " &
                        "selected_project_id = excluded.selected_project_id, " &
                        "current_page = excluded.current_page, " &
                        "search_text = excluded.search_text, " &
                        "active_filter = excluded.active_filter, " &
                        "language_code = excluded.language_code;"

                    command.Parameters.AddWithValue("$selectedProjectId", state.SelectedProjectId)
                    command.Parameters.AddWithValue("$currentPage", state.CurrentPage)
                    command.Parameters.AddWithValue("$searchText", state.SearchText)
                    command.Parameters.AddWithValue("$activeFilter", state.ActiveFilter)
                    command.Parameters.AddWithValue("$languageCode", state.LanguageCode)
                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Private Sub EnsureSchema()
            Using connection = CreateConnection()
                connection.Open()

                Using command = connection.CreateCommand()
                    command.CommandText =
                        "CREATE TABLE IF NOT EXISTS project_sessions (" &
                        "project_id TEXT NOT NULL PRIMARY KEY, " &
                        "project_name TEXT NOT NULL, " &
                        "is_archived INTEGER NOT NULL DEFAULT 0, " &
                        "project_summary TEXT NOT NULL DEFAULT '', " &
                        "selected_scenario_preset_id TEXT NOT NULL DEFAULT '', " &
                        "selected_planner_provider_key TEXT NOT NULL DEFAULT '', " &
                        "requirement_text TEXT NOT NULL, " &
                        "target_language TEXT NOT NULL, " &
                        "selected_planner_model TEXT NOT NULL, " &
                        "selected_image_provider_key TEXT NOT NULL DEFAULT '', " &
                        "selected_image_model TEXT NOT NULL DEFAULT '', " &
                        "selected_aspect_ratio TEXT NOT NULL, " &
                        "output_resolution TEXT NOT NULL, " &
                        "requested_count INTEGER NOT NULL, " &
                        "derivative_image_count INTEGER NOT NULL DEFAULT 1, " &
                        "auto_render_after_plan INTEGER NOT NULL DEFAULT 0, " &
                        "design_plan_markdown TEXT NOT NULL, " &
                        "image_plan_markdown TEXT NOT NULL, " &
                        "has_generated_plan INTEGER NOT NULL, " &
                        "last_output_directory TEXT NOT NULL, " &
                        "last_activated_at TEXT NOT NULL, " &
                        "uploaded_images_json TEXT NOT NULL, " &
                        "generated_images_json TEXT NOT NULL" &
                        ");" &
                        "CREATE TABLE IF NOT EXISTS scenario_presets (" &
                        "preset_id TEXT NOT NULL PRIMARY KEY, " &
                        "name TEXT NOT NULL, " &
                        "description TEXT NOT NULL DEFAULT '', " &
                        "planning_instruction TEXT NOT NULL DEFAULT '', " &
                        "image_instruction TEXT NOT NULL DEFAULT '', " &
                        "is_builtin INTEGER NOT NULL DEFAULT 0, " &
                        "sort_order INTEGER NOT NULL DEFAULT 0" &
                        ");" &
                        "CREATE TABLE IF NOT EXISTS studio_state (" &
                        "id INTEGER NOT NULL PRIMARY KEY CHECK(id = 1), " &
                        "selected_project_id TEXT NOT NULL, " &
                        "current_page INTEGER NOT NULL, " &
                        "search_text TEXT NOT NULL, " &
                        "active_filter INTEGER NOT NULL, " &
                        "language_code TEXT NOT NULL DEFAULT 'en'" &
                        ");"
                    command.ExecuteNonQuery()
                End Using

                EnsureColumn(connection, "project_sessions", "project_summary", "TEXT NOT NULL DEFAULT ''")
                EnsureColumn(connection, "project_sessions", "is_archived", "INTEGER NOT NULL DEFAULT 0")
                EnsureColumn(connection, "project_sessions", "selected_scenario_preset_id", "TEXT NOT NULL DEFAULT ''")
                EnsureColumn(connection, "project_sessions", "selected_planner_provider_key", "TEXT NOT NULL DEFAULT ''")
                EnsureColumn(connection, "project_sessions", "derivative_image_count", "INTEGER NOT NULL DEFAULT 1")
                EnsureColumn(connection, "project_sessions", "selected_image_provider_key", "TEXT NOT NULL DEFAULT ''")
                EnsureColumn(connection, "project_sessions", "selected_image_model", "TEXT NOT NULL DEFAULT ''")
                EnsureColumn(connection, "project_sessions", "auto_render_after_plan", "INTEGER NOT NULL DEFAULT 0")
                EnsureColumn(connection, "scenario_presets", "template_version", "INTEGER NOT NULL DEFAULT 1")
                EnsureColumn(connection, "studio_state", "language_code", "TEXT NOT NULL DEFAULT 'en'")
                EnsureScenarioSeedData(connection)
                UpgradeBuiltInScenarioTemplates(connection)
            End Using
        End Sub

        Private Function CreateConnection() As SqliteConnection
            Return New SqliteConnection($"Data Source={_databasePath}")
        End Function

        Private Shared Sub EnsureColumn(connection As SqliteConnection, tableName As String, columnName As String, definition As String)
            Using command = connection.CreateCommand()
                command.CommandText = $"PRAGMA table_info({tableName});"
                Using reader = command.ExecuteReader()
                    While reader.Read()
                        If String.Equals(reader.GetString(1), columnName, StringComparison.OrdinalIgnoreCase) Then
                            Return
                        End If
                    End While
                End Using
            End Using

            Using alterCommand = connection.CreateCommand()
                alterCommand.CommandText = $"ALTER TABLE {tableName} ADD COLUMN {columnName} {definition};"
                alterCommand.ExecuteNonQuery()
            End Using
        End Sub

        Private Shared Function DeserializeList(Of T)(json As String) As List(Of T)
            If String.IsNullOrWhiteSpace(json) Then
                Return New List(Of T)()
            End If

            Dim result = JsonSerializer.Deserialize(Of List(Of T))(json, JsonOptions)
            Return If(result, New List(Of T)())
        End Function

        Private Shared Function ParseRoundtripDateTime(value As String) As DateTime
            Dim parsed As DateTime
            If DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, parsed) Then
                Return parsed
            End If

            Return DateTime.Now
        End Function

        Private Shared Function BuildLines(ParamArray values() As String) As String
            Return String.Join(vbCrLf, values)
        End Function

        Private Shared Function GetDefaultEcommerceDesignPlanningTemplate() As String
            Return BuildLines(
                "你是一名资深电商详情页策划与视觉总监。请基于用户提供的产品图和需求，先完成产品分析，再输出统一的设计规划。",
                "要求：",
                "1. 必须先分析产品品类、结构、材质、卖点、适用人群和使用场景。",
                "2. 设计规划应完整描述产品总结、设计主题、目标受众、色彩体系、字体系统、视觉语言、摄影风格和版式指导。",
                "3. 整体结果要适合电商详情图生产，强调卖点层次、商业质感、版式可执行性和转化逻辑。",
                "4. 设计规划必须能够直接服务后续逐张图片规划。")
        End Function

        Private Shared Function GetDefaultEcommerceImagePlanningTemplate() As String
            Return BuildLines(
                "请基于整体设计规划继续输出逐张图片规划，最终图片数量必须与用户要求一致。",
                "要求：",
                "1. 每张图片规划都必须明确图片顺序、标题、目的、视角、场景、核心卖点、主标题、副标题和补充创意提示。",
                "2. 图片规划应覆盖主图展示、卖点拆解、细节特写、使用场景等电商常见信息层级。",
                "3. 如果目标语言为无文字，主标题和副标题必须留空。",
                "4. 如果目标语言不是无文字，主标题和副标题必须填写图片上实际要展示的文字内容。",
                "5. 生成的图片规划要便于直接生图，卖点清晰、分工明确、构图可执行。")
        End Function

        Private Shared Function GetDefaultTravelDesignPlanningTemplate() As String
            Return BuildLines(
                "你是一名资深旅行内容策划与视觉导演。请基于用户提供的图片和需求，先完成旅行主题分析，再输出统一的设计规划。",
                "要求：",
                "1. 必须先分析目的地特色、旅行方式、核心体验、人群偏好和适用场景。",
                "2. 设计规划应完整描述旅行主题、内容调性、受众画像、色彩氛围、字体风格、视觉语言、摄影风格和版式策略。",
                "3. 整体结果要强调目的地吸引力、故事线、代入感和出行灵感，适合旅行内容传播。",
                "4. 设计规划必须能够直接服务后续逐张图片规划。")
        End Function

        Private Shared Function GetDefaultTravelImagePlanningTemplate() As String
            Return BuildLines(
                "请基于整体设计规划继续输出逐张图片规划，最终图片数量必须与用户要求一致。",
                "要求：",
                "1. 每张图片规划都必须明确图片顺序、标题、目的、视角、场景、核心看点、主标题、副标题和补充创意提示。",
                "2. 图片规划应覆盖封面吸引图、目的地风光、路线体验、住宿美食、人物代入和行程亮点等内容层级。",
                "3. 如果目标语言为无文字，主标题和副标题必须留空。",
                "4. 如果目标语言不是无文字，主标题和副标题必须填写图片上实际要展示的文字内容。",
                "5. 生成的图片规划要便于直接生图，突出故事感、氛围感和旅行向往感。")
        End Function

        Private Shared Function GetDefaultXiaohongshuDesignPlanningTemplate() As String
            Return BuildLines(
                "你是一名资深小红书内容策划与视觉创意总监。请基于用户提供的图片和需求，先完成内容主题分析，再输出统一的设计规划。",
                "要求：",
                "1. 必须先分析内容主题、生活方式氛围、目标人群、种草卖点和分享场景。",
                "2. 设计规划应完整描述内容定位、封面策略、受众偏好、色彩气质、字体风格、视觉语言、摄影风格和版式节奏。",
                "3. 整体结果要强调真实感、审美统一、分享欲和收藏价值，避免过重广告感。",
                "4. 设计规划必须能够直接服务后续逐张图片规划。")
        End Function

        Private Shared Function GetDefaultXiaohongshuImagePlanningTemplate() As String
            Return BuildLines(
                "请基于整体设计规划继续输出逐张图片规划，最终图片数量必须与用户要求一致。",
                "要求：",
                "1. 每张图片规划都必须明确图片顺序、标题、目的、视角、场景、核心看点、主标题、副标题和补充创意提示。",
                "2. 图片规划应覆盖封面图、种草亮点、使用体验、细节氛围、生活方式场景和结尾总结图等内容层级。",
                "3. 如果目标语言为无文字，主标题和副标题必须留空。",
                "4. 如果目标语言不是无文字，主标题和副标题必须填写图片上实际要展示的文字内容。",
                "5. 生成的图片规划要便于直接生图，突出封面点击率、生活感、分享感和统一审美。")
        End Function

        Private Shared Sub EnsureScenarioSeedData(connection As SqliteConnection)
            Dim existingCount As Integer
            Using countCommand = connection.CreateCommand()
                countCommand.CommandText = "SELECT COUNT(*) FROM scenario_presets;"
                existingCount = Convert.ToInt32(countCommand.ExecuteScalar(), CultureInfo.InvariantCulture)
            End Using

            If existingCount > 0 Then
                Return
            End If

            Dim defaults = GetDefaultScenarioPresetRecords()
            For Each record In defaults
                Using insertCommand = connection.CreateCommand()
                    insertCommand.CommandText =
                        "INSERT INTO scenario_presets (preset_id, name, description, planning_instruction, image_instruction, is_builtin, sort_order, template_version) " &
                        "VALUES ($presetId, $name, $description, $planningInstruction, $imageInstruction, $isBuiltIn, $sortOrder, $templateVersion);"
                    insertCommand.Parameters.AddWithValue("$presetId", record.Id)
                    insertCommand.Parameters.AddWithValue("$name", record.Name)
                    insertCommand.Parameters.AddWithValue("$description", record.Description)
                    insertCommand.Parameters.AddWithValue("$planningInstruction", record.DesignPlanningTemplate)
                    insertCommand.Parameters.AddWithValue("$imageInstruction", record.ImagePlanningTemplate)
                    insertCommand.Parameters.AddWithValue("$isBuiltIn", If(record.IsBuiltIn, 1, 0))
                    insertCommand.Parameters.AddWithValue("$sortOrder", record.SortOrder)
                    insertCommand.Parameters.AddWithValue("$templateVersion", BuiltInScenarioTemplateVersion)
                    insertCommand.ExecuteNonQuery()
                End Using
            Next
        End Sub

        Private Shared Sub UpgradeBuiltInScenarioTemplates(connection As SqliteConnection)
            UpdateBuiltInPreset(
                connection,
                "preset_ecommerce",
                "E-commerce",
                "Suitable for product detail pages, marketplace hero images, and selling-point visuals, with emphasis on conversion flow, benefit hierarchy, and commercial polish.",
                GetDefaultEcommerceDesignPlanningTemplate(),
                GetDefaultEcommerceImagePlanningTemplate())

            UpdateBuiltInPreset(
                connection,
                "preset_travel",
                "旅行场景",
                "适合目的地内容、行程展示和旅行氛围表达，强调代入感、故事线和出行灵感。",
                GetDefaultTravelDesignPlanningTemplate(),
                GetDefaultTravelImagePlanningTemplate())

            UpdateBuiltInPreset(
                connection,
                "preset_xiaohongshu",
                "小红书场景",
                "适合种草笔记、封面图和生活方式内容，强调真实感、分享感和审美统一。",
                GetDefaultXiaohongshuDesignPlanningTemplate(),
                GetDefaultXiaohongshuImagePlanningTemplate())
        End Sub

        Private Shared Sub UpdateBuiltInPreset(connection As SqliteConnection,
                                               presetId As String,
                                               name As String,
                                               description As String,
                                               designPlanningTemplate As String,
                                               imagePlanningTemplate As String)
            Using command = connection.CreateCommand()
                command.CommandText =
                    "UPDATE scenario_presets " &
                    "SET name = $name, " &
                    "description = $description, " &
                    "planning_instruction = $planningInstruction, " &
                    "image_instruction = $imageInstruction, " &
                    "template_version = $templateVersion " &
                    "WHERE preset_id = $presetId AND is_builtin = 1 AND IFNULL(template_version, 1) < $templateVersion;"
                command.Parameters.AddWithValue("$presetId", presetId)
                command.Parameters.AddWithValue("$name", name)
                command.Parameters.AddWithValue("$description", description)
                command.Parameters.AddWithValue("$planningInstruction", designPlanningTemplate)
                command.Parameters.AddWithValue("$imageInstruction", imagePlanningTemplate)
                command.Parameters.AddWithValue("$templateVersion", BuiltInScenarioTemplateVersion)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Shared Function GetDefaultScenarioPresetRecords() As List(Of PersistedScenarioPresetRecord)
            Return New List(Of PersistedScenarioPresetRecord) From {
                New PersistedScenarioPresetRecord With {
                    .Id = "preset_ecommerce",
                    .Name = "E-commerce",
                    .Description = "Suitable for product detail pages, marketplace hero images, and selling-point visuals, with emphasis on conversion flow, benefit hierarchy, and commercial polish.",
                    .DesignPlanningTemplate = GetDefaultEcommerceDesignPlanningTemplate(),
                    .ImagePlanningTemplate = GetDefaultEcommerceImagePlanningTemplate(),
                    .IsBuiltIn = True,
                    .SortOrder = 1
                },
                New PersistedScenarioPresetRecord With {
                    .Id = "preset_travel",
                    .Name = "旅行场景",
                    .Description = "适合目的地内容、行程展示和旅行氛围表达，强调代入感、故事线和出行灵感。",
                    .DesignPlanningTemplate = GetDefaultTravelDesignPlanningTemplate(),
                    .ImagePlanningTemplate = GetDefaultTravelImagePlanningTemplate(),
                    .IsBuiltIn = True,
                    .SortOrder = 2
                },
                New PersistedScenarioPresetRecord With {
                    .Id = "preset_xiaohongshu",
                    .Name = "小红书场景",
                    .Description = "适合种草笔记、封面图和生活方式内容，强调真实感、分享感和审美统一。",
                    .DesignPlanningTemplate = GetDefaultXiaohongshuDesignPlanningTemplate(),
                    .ImagePlanningTemplate = GetDefaultXiaohongshuImagePlanningTemplate(),
                    .IsBuiltIn = True,
                    .SortOrder = 3
                }
            }
        End Function
    End Class
End Namespace

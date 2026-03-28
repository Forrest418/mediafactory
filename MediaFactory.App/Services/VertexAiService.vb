Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Text.Json
Imports System.Text.Json.Nodes
Imports System.Text.Json.Serialization.Metadata
Imports System.Windows.Media.Imaging
Imports Google.Apis.Auth.OAuth2
Imports MediaFactory.Models
Imports MediaFactory.Utils

Namespace Services
    Public Class VertexAiService
        Private Const CloudPlatformScope As String = "https://www.googleapis.com/auth/cloud-platform"
        Private Const PreferredLocalConfigFileName As String = "ModelProviders.local.json"
        Private Const PreferredConfigFileName As String = "ModelProviders.json"
        Private Const LegacyLocalConfigFileName As String = "GoogleVertex.local.json"
        Private Const LegacyConfigFileName As String = "GoogleVertex.json"
        Private Const DefaultGeminiImageModel As String = "gemini-3.1-flash-image-preview"
        Private Const DefaultGeminiApiBaseUrl As String = "https://generativelanguage.googleapis.com/v1beta"
        Private Const DefaultDeepSeekBaseUrl As String = "https://api.deepseek.com"
        Private Const DefaultDeepSeekModel As String = "deepseek-chat"
        Private Const DefaultOpenAiBaseUrl As String = "https://api.openai.com/v1"
        Private Const DefaultOpenAiPlannerModel As String = "gpt-4.1-mini"
        Private Const DefaultOpenAiImageModel As String = "gpt-image-1.5"
        Private Const DefaultQwenBaseUrl As String = "https://dashscope.aliyuncs.com/compatible-mode/v1"
        Private Const DefaultQwenPlannerModel As String = "qwen3.5-vl-plus"
        Private Const DefaultQwenImageEndpoint As String = "https://dashscope.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation"
        Private Const DefaultQwenImageModel As String = "qwen-image-2.0-pro"

        Private Shared ReadOnly HttpClientInstance As New HttpClient() With {
            .Timeout = TimeSpan.FromMinutes(5)
        }

        Private Shared ReadOnly JsonOptions As New JsonSerializerOptions With {
            .WriteIndented = False,
            .TypeInfoResolver = New DefaultJsonTypeInfoResolver()
        }

        Private ReadOnly _configPath As String
        Private ReadOnly _locationOverride As String
        Private _projectId As String
        Private _location As String
        Private _serviceAccountJson As String
        Private _defaultImageModel As String
        Private _plannerProviders As Dictionary(Of String, PlannerProviderSettings)
        Private _imageProviders As Dictionary(Of String, ImageProviderSettings)
        Private _plannerProvider As PlannerProviderSettings
        Private _imageProvider As ImageProviderSettings

        Public Sub New(Optional location As String = "global")
            _locationOverride = location
            _configPath = ResolveConfigPath()
            ReloadConfiguration()
        End Sub

        Public ReadOnly Property ProjectId As String
            Get
                Return _projectId
            End Get
        End Property

        Public ReadOnly Property Location As String
            Get
                Return _location
            End Get
        End Property

        Public ReadOnly Property PlannerProviderName As String
            Get
                If _plannerProvider Is Nothing Then
                    Return "未配置"
                End If

                Return _plannerProvider.ProviderName
            End Get
        End Property

        Public ReadOnly Property PlannerModelName As String
            Get
                If _plannerProvider Is Nothing Then
                    Return "未配置"
                End If

                Return _plannerProvider.Model
            End Get
        End Property

        Public ReadOnly Property PlannerDisplayName As String
            Get
                If _plannerProvider Is Nothing Then
                    Return "未配置"
                End If

                Return $"{_plannerProvider.ProviderName} / {_plannerProvider.Model}"
            End Get
        End Property

        Public ReadOnly Property PlannerProviderKey As String
            Get
                If _plannerProvider Is Nothing Then
                    Return String.Empty
                End If

                Return _plannerProvider.ProviderKey
            End Get
        End Property

        Public ReadOnly Property ImageProviderName As String
            Get
                If _imageProvider Is Nothing Then
                    Return "未配置"
                End If

                Return _imageProvider.ProviderName
            End Get
        End Property

        Public ReadOnly Property ImageProviderKey As String
            Get
                If _imageProvider Is Nothing Then
                    Return String.Empty
                End If

                Return _imageProvider.ProviderKey
            End Get
        End Property

        Public ReadOnly Property ImageModelName As String
            Get
                If _imageProvider Is Nothing Then
                    Return "未配置"
                End If

                Return _imageProvider.Model
            End Get
        End Property

        Public ReadOnly Property ImageDisplayName As String
            Get
                If _imageProvider Is Nothing Then
                    Return "未配置"
                End If

                Return $"{_imageProvider.ProviderName} / {_imageProvider.Model}"
            End Get
        End Property

        Public ReadOnly Property UsesCustomImageSize As Boolean
            Get
                Return _imageProvider IsNot Nothing AndAlso _imageProvider.RequiresFlexibleSize
            End Get
        End Property

        Public Function UsesCustomImageSizeForProvider(Optional imageProviderKey As String = Nothing) As Boolean
            Dim provider = ResolveRuntimeImageProvider(imageProviderKey, Nothing)
            Return provider IsNot Nothing AndAlso provider.RequiresFlexibleSize
        End Function

        Public Function GetDefaultPlannerProviderKey() As String
            Return If(_plannerProvider?.ProviderKey, "deepseek")
        End Function

        Public Function GetDefaultImageProviderKey() As String
            Return If(_imageProvider?.ProviderKey, "gemini")
        End Function

        Public Function GetDefaultPlannerModel(providerKey As String) As String
            Return ResolveRuntimePlannerProvider(providerKey, Nothing)?.Model
        End Function

        Public Function GetDefaultImageModel(providerKey As String) As String
            Return ResolveRuntimeImageProvider(providerKey, Nothing)?.Model
        End Function

        Public Function GetPlannerDisplayName(providerKey As String, plannerModel As String) As String
            Dim provider = ResolveRuntimePlannerProvider(providerKey, plannerModel)
            If provider Is Nothing Then
                Return "未配置"
            End If

            Return $"{provider.ProviderName} / {provider.Model}"
        End Function

        Public Function GetImageDisplayName(providerKey As String, imageModel As String) As String
            Dim provider = ResolveRuntimeImageProvider(providerKey, imageModel)
            If provider Is Nothing Then
                Return "未配置"
            End If

            Return $"{provider.ProviderName} / {provider.Model}"
        End Function

        Public Shared Function LoadModelConfiguration() As ModelProviderConfiguration
            Dim configPath = ResolveConfigPath()
            Dim root = JsonNode.Parse(File.ReadAllText(configPath, Encoding.UTF8))?.AsObject()
            If root Is Nothing Then
                Throw New InvalidOperationException($"{Path.GetFileName(configPath)} 不是有效的 JSON。")
            End If

            Return ParseModelConfiguration(root)
        End Function

        Public Shared Function LoadModelProviderEditors() As ModelProviderEditorState
            Dim configPath = ResolveConfigPath()
            Dim root = JsonNode.Parse(File.ReadAllText(configPath, Encoding.UTF8))?.AsObject()
            If root Is Nothing Then
                Throw New InvalidOperationException($"{Path.GetFileName(configPath)} 不是有效的 JSON。")
            End If

            Return New ModelProviderEditorState With {
                .PlannerProviders = BuildModelProviderEditors("planner", TryCast(root("planner"), JsonObject)),
                .ImageProviders = BuildModelProviderEditors("image", TryCast(root("image"), JsonObject))
            }
        End Function

        Public Shared Sub SaveModelProviderEditors(plannerProviders As IEnumerable(Of ModelProviderEditor),
                                                   imageProviders As IEnumerable(Of ModelProviderEditor))
            Dim configPath = ResolveConfigPath()
            Dim root = JsonNode.Parse(File.ReadAllText(configPath, Encoding.UTF8))?.AsObject()
            If root Is Nothing Then
                Throw New InvalidOperationException($"{Path.GetFileName(configPath)} 不是有效的 JSON。")
            End If

            ApplyModelProviderEditors(EnsureObject(root, "planner"), plannerProviders)
            ApplyModelProviderEditors(EnsureObject(root, "image"), imageProviders)

            File.WriteAllText(configPath, root.ToJsonString(New JsonSerializerOptions With {
                .WriteIndented = True,
                .TypeInfoResolver = New DefaultJsonTypeInfoResolver()
            }), Encoding.UTF8)
        End Sub

        Public Shared Sub SaveModelConfiguration(settings As ModelProviderConfiguration)
            If settings Is Nothing Then
                Throw New ArgumentNullException(NameOf(settings))
            End If

            Dim configPath = ResolveConfigPath()
            Dim root = JsonNode.Parse(File.ReadAllText(configPath, Encoding.UTF8))?.AsObject()
            If root Is Nothing Then
                Throw New InvalidOperationException($"{Path.GetFileName(configPath)} 不是有效的 JSON。")
            End If

            Dim plannerObject = EnsureObject(root, "planner")
            plannerObject("selected_provider") = settings.SelectedPlannerProviderKey

            Dim deepseekObject = EnsureObject(plannerObject, "deepseek")
            deepseekObject("api_key") = settings.DeepSeekApiKey
            deepseekObject("base_url") = settings.DeepSeekBaseUrl
            deepseekObject("model") = settings.DeepSeekModel

            Dim openAiPlannerObject = EnsureObject(plannerObject, "openai")
            openAiPlannerObject("api_key") = settings.OpenAiPlannerApiKey
            openAiPlannerObject("base_url") = settings.OpenAiPlannerBaseUrl
            openAiPlannerObject("model") = settings.OpenAiPlannerModel

            Dim qwenPlannerObject = EnsureObject(plannerObject, "qwen")
            qwenPlannerObject("api_key") = settings.QwenPlannerApiKey
            qwenPlannerObject("base_url") = settings.QwenPlannerBaseUrl
            qwenPlannerObject("model") = settings.QwenPlannerModel

            Dim imageObject = EnsureObject(root, "image")
            imageObject("selected_provider") = settings.SelectedImageProviderKey

            Dim openAiImageObject = EnsureObject(imageObject, "openai")
            openAiImageObject("api_key") = settings.OpenAiImageApiKey
            openAiImageObject("base_url") = settings.OpenAiImageBaseUrl
            openAiImageObject("model") = settings.OpenAiImageModel

            Dim qwenImageObject = EnsureObject(imageObject, "qwen")
            qwenImageObject("api_key") = settings.QwenImageApiKey
            qwenImageObject("base_url") = settings.QwenImageBaseUrl
            qwenImageObject("model") = settings.QwenImageModel

            Dim geminiObject = EnsureObject(imageObject, "gemini")
            geminiObject("api_key") = settings.GeminiApiKey
            geminiObject("model") = settings.GeminiImageModel

            Dim vertexObject = EnsureObject(root, "vertex")
            vertexObject("image_model") = settings.GeminiImageModel

            File.WriteAllText(configPath, root.ToJsonString(New JsonSerializerOptions With {
                .WriteIndented = True,
                .TypeInfoResolver = New DefaultJsonTypeInfoResolver()
            }), Encoding.UTF8)
        End Sub

        Public Sub ReloadConfiguration()
            Dim root = JsonNode.Parse(File.ReadAllText(_configPath, Encoding.UTF8))?.AsObject()
            If root Is Nothing Then
                Throw New InvalidOperationException($"{Path.GetFileName(_configPath)} 不是有效的 JSON。")
            End If

            If root.ContainsKey("vertex") OrElse root.ContainsKey("planner") OrElse root.ContainsKey("image") Then
                LoadFromNestedConfiguration(root, _locationOverride)
            Else
                LoadFromLegacyConfiguration(root, _locationOverride)
            End If
        End Sub

        Public Async Function AnalyzeProductAsync(images As IEnumerable(Of ProductImageItem),
                                                  request As AnalyzeRequest,
                                                  Optional plannerProviderKey As String = Nothing,
                                                  Optional plannerModel As String = Nothing,
                                                  Optional cancellationToken As Threading.CancellationToken = Nothing,
                                                  Optional log As Action(Of String) = Nothing) As Task(Of DesignPlan)
            Dim activePlannerProvider = ResolveRuntimePlannerProvider(plannerProviderKey, plannerModel)

            If activePlannerProvider Is Nothing Then
                Throw New InvalidOperationException($"未在 {Path.GetFileName(_configPath)} 中配置规划模型。请先在系统设置中完成规划模型配置。")
            End If

            If String.IsNullOrWhiteSpace(activePlannerProvider.ApiKey) Then
                Throw New InvalidOperationException($"当前规划模型 {activePlannerProvider.ProviderName} / {activePlannerProvider.Model} 缺少 API Key，请先到系统设置中补全。")
            End If

            Dim safePlannerModel = activePlannerProvider.Model
            log?.Invoke($"规划请求准备完成：provider={activePlannerProvider.ProviderName}, model={safePlannerModel}, supportsVision={activePlannerProvider.SupportsVision}, imageCount={images.Count()}, requestedCount={request.RequestedCount}, language={request.TargetLanguage}, aspectRatio={request.AspectRatio}")

            Dim responseText = Await RequestPlannerCompletionAsync(images, request, activePlannerProvider, safePlannerModel, cancellationToken, log)

            If String.IsNullOrWhiteSpace(responseText) Then
                Throw New InvalidOperationException($"{activePlannerProvider.ProviderName} 没有返回可解析的规划结果。")
            End If

            log?.Invoke($"规划请求成功：provider={activePlannerProvider.ProviderName}, model={safePlannerModel}, responseLength={responseText.Length}")
            Return ParseDesignPlan(responseText)
        End Function

        Public Async Function GenerateImagesAsync(images As IEnumerable(Of ProductImageItem),
                                                  imagePlans As IEnumerable(Of ImagePlanItem),
                                                  imageProviderKey As String,
                                                  imageModel As String,
                                                  aspectRatio As String,
                                                  imageSize As String,
                                                  targetLanguage As String,
                                                  scenarioName As String,
                                                  scenarioImageInstruction As String,
                                                  outputDirectory As String,
                                                  progress As IProgress(Of String),
                                                  Optional cancellationToken As Threading.CancellationToken = Nothing,
                                                  Optional log As Action(Of String) = Nothing) As Task(Of List(Of GeneratedImageItem))
            Directory.CreateDirectory(outputDirectory)

            Dim plans = imagePlans.ToList()
            Dim results As New List(Of GeneratedImageItem)()
            Dim activeImageProvider = ResolveRuntimeImageProvider(imageProviderKey, imageModel)
            Dim safeModel = activeImageProvider.Model

            log?.Invoke($"图片生成批次开始：provider={activeImageProvider.ProviderName}, model={safeModel}, scenario={scenarioName}, aspectRatio={aspectRatio}, imageSize={imageSize}, targetLanguage={targetLanguage}, outputDirectory={outputDirectory}, sourceImageCount={images.Count()}, planCount={plans.Count}")

            For index = 0 To plans.Count - 1
                cancellationToken.ThrowIfCancellationRequested()
                Dim plan = plans(index)
                progress?.Report(String.Format(CultureInfo.InvariantCulture, "正在生成第 {0}/{1} 张：{2}", index + 1, plans.Count, plan.Title))
                log?.Invoke($"开始生成图{plan.Sequence}：title={plan.Title}, scene={plan.Scene}, sellingPoint={plan.SellingPoint}")

                Dim imageData = Await GenerateSingleImageBytesAsync(
                    activeImageProvider,
                    images,
                    plan,
                    safeModel,
                    aspectRatio,
                    imageSize,
                    targetLanguage,
                    scenarioName,
                    scenarioImageInstruction,
                    cancellationToken,
                    log)

                If imageData Is Nothing OrElse imageData.Length = 0 Then
                    Throw New InvalidOperationException($"模型未返回图片数据，失败项：{plan.Title}")
                End If

                Dim fileName = $"{plan.Sequence:00}_{SanitizeFileName(plan.Title)}.png"
                Dim savedPath = Path.Combine(outputDirectory, fileName)
                Await File.WriteAllBytesAsync(savedPath, imageData, cancellationToken)

                results.Add(New GeneratedImageItem With {
                    .Title = plan.Title,
                    .Description = plan.Purpose,
                    .SavedPath = savedPath,
                    .Preview = LoadBitmap(savedPath)
                })
                log?.Invoke($"图{plan.Sequence} 生成成功：savedPath={savedPath}")
            Next

            Return results
        End Function

        Private Sub LoadFromNestedConfiguration(root As JsonObject, locationOverride As String)
            Dim plannerObject = TryCast(root("planner"), JsonObject)
            Dim imageObject = TryCast(root("image"), JsonObject)
            _plannerProviders = BuildPlannerProviders(plannerObject)
            _plannerProvider = ResolvePlannerProvider(plannerObject)
            _imageProviders = BuildImageProviders(root, plannerObject, imageObject)
            _imageProvider = ResolveImageProvider(root, plannerObject, imageObject)

            Dim vertexObject = TryCast(root("vertex"), JsonObject)
            _projectId = ReadString(vertexObject, "project_id")
            _defaultImageModel = ReadString(vertexObject, "image_model", ReadString(TryCast(GetChildObject(imageObject, "gemini"), JsonObject), "model", DefaultGeminiImageModel))
            _location = ResolveLocation(ReadString(vertexObject, "location"), locationOverride)
            _serviceAccountJson = String.Empty

            Dim serviceAccountObject As JsonObject = Nothing
            If vertexObject IsNot Nothing Then
                serviceAccountObject = TryCast(vertexObject("service_account"), JsonObject)
            End If

            If serviceAccountObject Is Nothing Then
                Return
            End If

            _serviceAccountJson = serviceAccountObject.ToJsonString(JsonOptions)

            If String.IsNullOrWhiteSpace(_projectId) Then
                _projectId = ReadString(serviceAccountObject, "project_id")
            End If
        End Sub

        Private Sub LoadFromLegacyConfiguration(root As JsonObject, locationOverride As String)
            _plannerProviders = New Dictionary(Of String, PlannerProviderSettings)(StringComparer.OrdinalIgnoreCase)
            _plannerProvider = Nothing
            _imageProviders = New Dictionary(Of String, ImageProviderSettings)(StringComparer.OrdinalIgnoreCase)
            _imageProvider = New ImageProviderSettings With {
                .ProviderKey = "gemini",
                .ProviderName = "Gemini",
                .ApiKey = String.Empty,
                .BaseUrl = String.Empty,
                .Model = DefaultGeminiImageModel,
                .RequiresFlexibleSize = False,
                .UsesVertex = True
            }
            _imageProviders("gemini") = CloneImageProvider(_imageProvider)
            _serviceAccountJson = root.ToJsonString(JsonOptions)
            _projectId = ReadString(root, "project_id")
            _defaultImageModel = DefaultGeminiImageModel
            _location = ResolveLocation(Nothing, locationOverride)

            If String.IsNullOrWhiteSpace(_projectId) Then
                Throw New InvalidOperationException($"{Path.GetFileName(_configPath)} 中缺少 project_id，无法调用 Vertex AI。")
            End If
        End Sub

        Private Function ResolvePlannerProvider(plannerObject As JsonObject) As PlannerProviderSettings
            If plannerObject Is Nothing Then
                Return Nothing
            End If

            Dim selectedProvider = ReadString(plannerObject, "selected_provider").ToLowerInvariant()
            If String.Equals(selectedProvider, "openai", StringComparison.OrdinalIgnoreCase) Then
                Return BuildPlannerProvider(
                    providerKey:="openai",
                    providerName:="OpenAI Compatible",
                    providerObject:=TryCast(plannerObject("openai"), JsonObject),
                    defaultBaseUrl:=DefaultOpenAiBaseUrl,
                    defaultModel:=DefaultOpenAiPlannerModel,
                    supportsVision:=True)
            End If

            If String.Equals(selectedProvider, "qwen", StringComparison.OrdinalIgnoreCase) Then
                Return BuildPlannerProvider(
                    providerKey:="qwen",
                    providerName:="Qwen",
                    providerObject:=TryCast(plannerObject("qwen"), JsonObject),
                    defaultBaseUrl:=DefaultQwenBaseUrl,
                    defaultModel:=DefaultQwenPlannerModel,
                    supportsVision:=True)
            End If

            If String.Equals(selectedProvider, "deepseek", StringComparison.OrdinalIgnoreCase) Then
                Return BuildPlannerProvider(
                    providerKey:="deepseek",
                    providerName:="DeepSeek",
                    providerObject:=TryCast(plannerObject("deepseek"), JsonObject),
                    defaultBaseUrl:=DefaultDeepSeekBaseUrl,
                    defaultModel:=DefaultDeepSeekModel,
                    supportsVision:=False)
            End If

            Dim deepSeek = BuildPlannerProvider(
                providerKey:="deepseek",
                providerName:="DeepSeek",
                providerObject:=TryCast(plannerObject("deepseek"), JsonObject),
                defaultBaseUrl:=DefaultDeepSeekBaseUrl,
                defaultModel:=DefaultDeepSeekModel,
                supportsVision:=False)

            If deepSeek IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(deepSeek.ApiKey) Then
                Return deepSeek
            End If

            Dim openAiPlanner = BuildPlannerProvider(
                providerKey:="openai",
                providerName:="OpenAI Compatible",
                providerObject:=TryCast(plannerObject("openai"), JsonObject),
                defaultBaseUrl:=DefaultOpenAiBaseUrl,
                defaultModel:=DefaultOpenAiPlannerModel,
                supportsVision:=True)

            If openAiPlanner IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(openAiPlanner.ApiKey) Then
                Return openAiPlanner
            End If

            Return BuildPlannerProvider(
                providerKey:="qwen",
                providerName:="Qwen",
                providerObject:=TryCast(plannerObject("qwen"), JsonObject),
                defaultBaseUrl:=DefaultQwenBaseUrl,
                defaultModel:=DefaultQwenPlannerModel,
                supportsVision:=True)
        End Function

        Private Function BuildPlannerProviders(plannerObject As JsonObject) As Dictionary(Of String, PlannerProviderSettings)
            Dim providers As New Dictionary(Of String, PlannerProviderSettings)(StringComparer.OrdinalIgnoreCase)
            providers("deepseek") = BuildPlannerProvider(
                providerKey:="deepseek",
                providerName:="DeepSeek",
                providerObject:=TryCast(GetChildObject(plannerObject, "deepseek"), JsonObject),
                defaultBaseUrl:=DefaultDeepSeekBaseUrl,
                defaultModel:=DefaultDeepSeekModel,
                supportsVision:=False)
            providers("openai") = BuildPlannerProvider(
                providerKey:="openai",
                providerName:="OpenAI Compatible",
                providerObject:=TryCast(GetChildObject(plannerObject, "openai"), JsonObject),
                defaultBaseUrl:=DefaultOpenAiBaseUrl,
                defaultModel:=DefaultOpenAiPlannerModel,
                supportsVision:=True)
            providers("qwen") = BuildPlannerProvider(
                providerKey:="qwen",
                providerName:="Qwen",
                providerObject:=TryCast(GetChildObject(plannerObject, "qwen"), JsonObject),
                defaultBaseUrl:=DefaultQwenBaseUrl,
                defaultModel:=DefaultQwenPlannerModel,
                supportsVision:=True)
            Return providers
        End Function

        Private Function ResolveImageProvider(root As JsonObject,
                                              plannerObject As JsonObject,
                                              imageObject As JsonObject) As ImageProviderSettings
            Dim selectedProvider = ReadString(imageObject, "selected_provider")
            Dim openAiImageProvider = BuildOpenAiImageProvider(imageObject)
            Dim qwenImageObject = TryCast(GetChildObject(imageObject, "qwen"), JsonObject)
            Dim plannerQwenObject = TryCast(GetChildObject(plannerObject, "qwen"), JsonObject)

            Dim qwenImageProvider = New ImageProviderSettings With {
                .ProviderKey = "qwen",
                .ProviderName = "Qwen",
                .ApiKey = ReadString(qwenImageObject, "api_key", ReadString(plannerQwenObject, "api_key")),
                .BaseUrl = NormalizeEndpoint(ReadString(qwenImageObject, "base_url", DefaultQwenImageEndpoint), DefaultQwenImageEndpoint),
                .Model = ReadString(qwenImageObject, "model", DefaultQwenImageModel),
                .RequiresFlexibleSize = True,
                .UsesVertex = False
            }

            Dim geminiProvider = BuildGeminiImageProvider(root, imageObject)

            If String.Equals(selectedProvider, "openai", StringComparison.OrdinalIgnoreCase) Then
                Return openAiImageProvider
            End If

            If String.Equals(selectedProvider, "qwen", StringComparison.OrdinalIgnoreCase) Then
                Return qwenImageProvider
            End If

            If String.Equals(selectedProvider, "gemini", StringComparison.OrdinalIgnoreCase) Then
                Return geminiProvider
            End If

            If Not String.IsNullOrWhiteSpace(openAiImageProvider.ApiKey) Then
                Return openAiImageProvider
            End If

            If Not String.IsNullOrWhiteSpace(qwenImageProvider.ApiKey) AndAlso
               String.IsNullOrWhiteSpace(geminiProvider.ApiKey) AndAlso
               Not HasVertexServiceAccountConfiguration(TryCast(root("vertex"), JsonObject)) Then
                Return qwenImageProvider
            End If

            Return geminiProvider
        End Function

        Private Function BuildImageProviders(root As JsonObject,
                                             plannerObject As JsonObject,
                                             imageObject As JsonObject) As Dictionary(Of String, ImageProviderSettings)
            Dim providers As New Dictionary(Of String, ImageProviderSettings)(StringComparer.OrdinalIgnoreCase)
            Dim qwenImageObject = TryCast(GetChildObject(imageObject, "qwen"), JsonObject)
            Dim plannerQwenObject = TryCast(GetChildObject(plannerObject, "qwen"), JsonObject)

            providers("openai") = BuildOpenAiImageProvider(imageObject)
            providers("qwen") = New ImageProviderSettings With {
                .ProviderKey = "qwen",
                .ProviderName = "Qwen",
                .ApiKey = ReadString(qwenImageObject, "api_key", ReadString(plannerQwenObject, "api_key")),
                .BaseUrl = NormalizeEndpoint(ReadString(qwenImageObject, "base_url", DefaultQwenImageEndpoint), DefaultQwenImageEndpoint),
                .Model = ReadString(qwenImageObject, "model", DefaultQwenImageModel),
                .RequiresFlexibleSize = True,
                .UsesVertex = False
            }

            providers("gemini") = BuildGeminiImageProvider(root, imageObject)

            Return providers
        End Function

        Private Shared Function BuildOpenAiImageProvider(imageObject As JsonObject) As ImageProviderSettings
            Dim openAiObject = TryCast(GetChildObject(imageObject, "openai"), JsonObject)

            Return New ImageProviderSettings With {
                .ProviderKey = "openai",
                .ProviderName = "OpenAI Compatible",
                .ApiKey = ReadString(openAiObject, "api_key"),
                .BaseUrl = NormalizeBaseUrl(ReadString(openAiObject, "base_url", DefaultOpenAiBaseUrl)),
                .Model = ReadString(openAiObject, "model", DefaultOpenAiImageModel),
                .RequiresFlexibleSize = False,
                .UsesVertex = False
            }
        End Function

        Private Shared Function BuildGeminiImageProvider(root As JsonObject,
                                                         imageObject As JsonObject) As ImageProviderSettings
            Dim geminiObject = TryCast(GetChildObject(imageObject, "gemini"), JsonObject)
            Dim geminiApiKey = ReadString(geminiObject, "api_key")

            Return New ImageProviderSettings With {
                .ProviderKey = "gemini",
                .ProviderName = "Gemini",
                .ApiKey = geminiApiKey,
                .BaseUrl = DefaultGeminiApiBaseUrl,
                .Model = ReadString(geminiObject, "model", ReadString(TryCast(root("vertex"), JsonObject), "image_model", DefaultGeminiImageModel)),
                .RequiresFlexibleSize = False,
                .UsesVertex = String.IsNullOrWhiteSpace(geminiApiKey)
            }
        End Function

        Private Shared Function BuildPlannerProvider(providerKey As String,
                                                     providerName As String,
                                                     providerObject As JsonObject,
                                                     defaultBaseUrl As String,
                                                     defaultModel As String,
                                                     supportsVision As Boolean) As PlannerProviderSettings
            If providerObject Is Nothing Then
                providerObject = New JsonObject()
            End If

            Return New PlannerProviderSettings With {
                .ProviderKey = providerKey,
                .ProviderName = providerName,
                .ApiKey = ReadString(providerObject, "api_key"),
                .BaseUrl = NormalizeBaseUrl(ReadString(providerObject, "base_url", defaultBaseUrl)),
                .Model = ReadString(providerObject, "model", defaultModel),
                .SupportsVision = supportsVision
            }
        End Function

        Private Function ResolveRuntimePlannerProvider(providerKey As String, modelName As String) As PlannerProviderSettings
            If _plannerProviders Is Nothing OrElse _plannerProviders.Count = 0 Then
                If _plannerProvider Is Nothing Then
                    Return Nothing
                End If

                Dim fallbackProvider = ClonePlannerProvider(_plannerProvider)
                If Not String.IsNullOrWhiteSpace(modelName) Then
                    fallbackProvider.Model = modelName.Trim()
                End If
                Return fallbackProvider
            End If

            Dim normalizedKey = If(String.IsNullOrWhiteSpace(providerKey), _plannerProvider?.ProviderKey, providerKey).Trim()
            If String.IsNullOrWhiteSpace(normalizedKey) OrElse Not _plannerProviders.ContainsKey(normalizedKey) Then
                normalizedKey = If(_plannerProvider?.ProviderKey, _plannerProviders.Keys.FirstOrDefault())
            End If

            If String.IsNullOrWhiteSpace(normalizedKey) Then
                Return Nothing
            End If

            Dim resolved = ClonePlannerProvider(_plannerProviders(normalizedKey))
            If Not String.IsNullOrWhiteSpace(modelName) Then
                resolved.Model = modelName.Trim()
            End If

            Return resolved
        End Function

        Private Function ResolveRuntimeImageProvider(providerKey As String, modelName As String) As ImageProviderSettings
            If _imageProviders Is Nothing OrElse _imageProviders.Count = 0 Then
                If _imageProvider Is Nothing Then
                    Throw New InvalidOperationException("当前未配置图片生成模型。请先到系统设置中完成配置。")
                End If

                Dim fallbackProvider = CloneImageProvider(_imageProvider)
                If Not String.IsNullOrWhiteSpace(modelName) Then
                    fallbackProvider.Model = modelName.Trim()
                End If
                Return fallbackProvider
            End If

            Dim normalizedKey = If(String.IsNullOrWhiteSpace(providerKey), _imageProvider?.ProviderKey, providerKey).Trim()
            If String.IsNullOrWhiteSpace(normalizedKey) OrElse Not _imageProviders.ContainsKey(normalizedKey) Then
                normalizedKey = If(_imageProvider?.ProviderKey, _imageProviders.Keys.FirstOrDefault())
            End If

            If String.IsNullOrWhiteSpace(normalizedKey) Then
                Throw New InvalidOperationException("当前未配置图片生成模型。请先到系统设置中完成配置。")
            End If

            Dim resolved = CloneImageProvider(_imageProviders(normalizedKey))
            If Not String.IsNullOrWhiteSpace(modelName) Then
                resolved.Model = modelName.Trim()
            End If

            If resolved.UsesVertex Then
                EnsureImageGenerationModelSupported(resolved.Model)
                EnsureVertexImageConfigurationAvailable(resolved)
            ElseIf String.IsNullOrWhiteSpace(resolved.ApiKey) Then
                Throw New InvalidOperationException($"当前绘图模型 {resolved.ProviderName} / {resolved.Model} 缺少 API Key，请先到系统设置中补全。")
            End If

            Return resolved
        End Function

        Private Async Function GenerateSingleImageBytesAsync(imageProvider As ImageProviderSettings,
                                                             images As IEnumerable(Of ProductImageItem),
                                                             plan As ImagePlanItem,
                                                             modelName As String,
                                                             aspectRatio As String,
                                                             imageSize As String,
                                                             targetLanguage As String,
                                                             scenarioName As String,
                                                             scenarioImageInstruction As String,
                                                             cancellationToken As Threading.CancellationToken,
                                                             log As Action(Of String)) As Task(Of Byte())
            If imageProvider.UsesVertex Then
                Return Await GenerateVertexImageBytesAsync(images, plan, modelName, aspectRatio, imageSize, targetLanguage, scenarioName, scenarioImageInstruction, cancellationToken, log)
            End If

            If String.Equals(imageProvider.ProviderKey, "gemini", StringComparison.OrdinalIgnoreCase) Then
                Return Await GenerateGeminiApiImageBytesAsync(images, plan, imageProvider, aspectRatio, imageSize, targetLanguage, scenarioName, scenarioImageInstruction, cancellationToken, log)
            End If

            If String.Equals(imageProvider.ProviderKey, "openai", StringComparison.OrdinalIgnoreCase) Then
                Return Await GenerateOpenAiImageBytesAsync(images, plan, imageProvider, aspectRatio, imageSize, targetLanguage, scenarioName, scenarioImageInstruction, cancellationToken, log)
            End If

            Return Await GenerateQwenImageBytesAsync(images, plan, imageProvider, imageSize, targetLanguage, scenarioName, scenarioImageInstruction, cancellationToken, log)
        End Function

        Private Async Function GenerateVertexImageBytesAsync(images As IEnumerable(Of ProductImageItem),
                                                             plan As ImagePlanItem,
                                                             modelName As String,
                                                             aspectRatio As String,
                                                             imageSize As String,
                                                             targetLanguage As String,
                                                             scenarioName As String,
                                                             scenarioImageInstruction As String,
                                                             cancellationToken As Threading.CancellationToken,
                                                             log As Action(Of String)) As Task(Of Byte())
            Dim parts As New JsonArray From {
                BuildTextPart(BuildImagePrompt(plan, targetLanguage, scenarioName, scenarioImageInstruction))
            }

            For Each item In images
                parts.Add(BuildImagePart(item.FullPath))
            Next

            Dim requestBody As New JsonObject From {
                {"contents", New JsonArray From {
                    New JsonObject From {
                        {"role", "user"},
                        {"parts", parts}
                    }
                }},
                {"generationConfig", New JsonObject From {
                    {"responseModalities", New JsonArray From {"TEXT", "IMAGE"}},
                    {"candidateCount", 1},
                    {"imageConfig", New JsonObject From {
                        {"aspectRatio", aspectRatio},
                        {"imageSize", imageSize}
                    }}
                }}
            }

            Dim responseJson = Await PostVertexGenerateContentAsync(modelName, requestBody, cancellationToken, log)
            Return ExtractImageData(responseJson)
        End Function

        Private Async Function GenerateGeminiApiImageBytesAsync(images As IEnumerable(Of ProductImageItem),
                                                                plan As ImagePlanItem,
                                                                imageProvider As ImageProviderSettings,
                                                                aspectRatio As String,
                                                                imageSize As String,
                                                                targetLanguage As String,
                                                                scenarioName As String,
                                                                scenarioImageInstruction As String,
                                                                cancellationToken As Threading.CancellationToken,
                                                                log As Action(Of String)) As Task(Of Byte())
            Dim parts As New JsonArray From {
                BuildTextPart(BuildImagePrompt(plan, targetLanguage, scenarioName, scenarioImageInstruction))
            }

            For Each item In images
                parts.Add(BuildGeminiApiImagePart(item.FullPath))
            Next

            Dim requestBody As New JsonObject From {
                {"contents", New JsonArray From {
                    New JsonObject From {
                        {"role", "user"},
                        {"parts", parts}
                    }
                }},
                {"generationConfig", New JsonObject From {
                    {"responseModalities", New JsonArray From {"TEXT", "IMAGE"}},
                    {"candidateCount", 1},
                    {"imageConfig", New JsonObject From {
                        {"aspectRatio", aspectRatio},
                        {"imageSize", imageSize}
                    }}
                }}
            }

            Dim responseJson = Await PostGeminiGenerateContentAsync(imageProvider, requestBody, cancellationToken, log)
            Return ExtractImageData(responseJson)
        End Function

        Private Async Function GenerateQwenImageBytesAsync(images As IEnumerable(Of ProductImageItem),
                                                           plan As ImagePlanItem,
                                                           imageProvider As ImageProviderSettings,
                                                           imageSize As String,
                                                           targetLanguage As String,
                                                           scenarioName As String,
                                                           scenarioImageInstruction As String,
                                                           cancellationToken As Threading.CancellationToken,
                                                           log As Action(Of String)) As Task(Of Byte())
            Dim safeSize = NormalizeQwenImageSize(imageSize)
            Dim sourceImages = images.Where(Function(item) item IsNot Nothing AndAlso File.Exists(item.FullPath)).Take(3).ToList()
            If sourceImages.Count = 0 Then
                Throw New InvalidOperationException("Qwen 绘图至少需要 1 张有效的产品图。")
            End If

            If images.Count() > sourceImages.Count Then
                log?.Invoke($"Qwen 出图仅支持最多 3 张参考图，已截取前 {sourceImages.Count} 张。")
            End If

            Dim content As New JsonArray()
            For Each item In sourceImages
                content.Add(New JsonObject From {
                    {"image", BuildDataUrl(item.FullPath)}
                })
            Next

            content.Add(New JsonObject From {
                {"text", BuildImagePrompt(plan, targetLanguage, scenarioName, scenarioImageInstruction)}
            })

            Dim requestBody As New JsonObject From {
                {"model", imageProvider.Model},
                {"input", New JsonObject From {
                    {"messages", New JsonArray From {
                        New JsonObject From {
                            {"role", "user"},
                            {"content", content}
                        }
                    }}
                }},
                {"parameters", New JsonObject From {
                    {"n", 1},
                    {"size", safeSize}
                }}
            }

            Dim responseJson = Await PostQwenImageGenerationAsync(imageProvider, requestBody, cancellationToken, log)
            Dim resultUrl = ExtractQwenImageUrl(responseJson)
            If String.IsNullOrWhiteSpace(resultUrl) Then
                Throw New InvalidOperationException("Qwen 未返回可下载的图片地址。")
            End If

            Return Await DownloadBinaryAsync(resultUrl, cancellationToken)
        End Function

        Private Async Function GenerateOpenAiImageBytesAsync(images As IEnumerable(Of ProductImageItem),
                                                             plan As ImagePlanItem,
                                                             imageProvider As ImageProviderSettings,
                                                             aspectRatio As String,
                                                             imageSize As String,
                                                             targetLanguage As String,
                                                             scenarioName As String,
                                                             scenarioImageInstruction As String,
                                                             cancellationToken As Threading.CancellationToken,
                                                             log As Action(Of String)) As Task(Of Byte())
            Dim sourceImages = images.Where(Function(item) item IsNot Nothing AndAlso File.Exists(item.FullPath)).ToList()
            If sourceImages.Count = 0 Then
                Throw New InvalidOperationException("OpenAI compatible image editing requires at least one valid reference image.")
            End If

            Dim requestBody As New MultipartFormDataContent()
            requestBody.Add(New StringContent(imageProvider.Model, Encoding.UTF8), "model")
            requestBody.Add(New StringContent(BuildImagePrompt(plan, targetLanguage, scenarioName, scenarioImageInstruction), Encoding.UTF8), "prompt")
            requestBody.Add(New StringContent("1", Encoding.UTF8), "n")
            requestBody.Add(New StringContent(MapOpenAiImageSize(aspectRatio, imageSize), Encoding.UTF8), "size")
            requestBody.Add(New StringContent("high", Encoding.UTF8), "input_fidelity")

            For Each item In sourceImages
                Dim imageContent As New ByteArrayContent(File.ReadAllBytes(item.FullPath))
                imageContent.Headers.ContentType = New MediaTypeHeaderValue(GetMimeTypeFromExtension(Path.GetExtension(item.FullPath)))
                requestBody.Add(imageContent, "image[]", Path.GetFileName(item.FullPath))
            Next

            Dim responseJson = Await PostOpenAiImageEditAsync(imageProvider, requestBody, cancellationToken, log)
            Return Await ExtractOpenAiImageBytesAsync(responseJson, cancellationToken)
        End Function

        Private Shared Function ResolveLocation(configuredLocation As String, locationOverride As String) As String
            If Not String.IsNullOrWhiteSpace(locationOverride) Then
                Return locationOverride.Trim()
            End If

            If String.IsNullOrWhiteSpace(configuredLocation) Then
                Return "global"
            End If

            Return configuredLocation.Trim()
        End Function

        Private Shared Function ParseModelConfiguration(root As JsonObject) As ModelProviderConfiguration
            Dim plannerObject = TryCast(root("planner"), JsonObject)
            Dim imageObject = TryCast(root("image"), JsonObject)
            Dim vertexObject = TryCast(root("vertex"), JsonObject)
            Dim plannerOpenAiObject = TryCast(GetChildObject(plannerObject, "openai"), JsonObject)
            Dim plannerQwenObject = TryCast(GetChildObject(plannerObject, "qwen"), JsonObject)
            Dim imageOpenAiObject = TryCast(GetChildObject(imageObject, "openai"), JsonObject)
            Dim imageQwenObject = TryCast(GetChildObject(imageObject, "qwen"), JsonObject)

            Return New ModelProviderConfiguration With {
                .SelectedPlannerProviderKey = ReadString(plannerObject, "selected_provider", ResolveDefaultPlannerProviderKey(plannerObject)),
                .DeepSeekApiKey = ReadString(TryCast(GetChildObject(plannerObject, "deepseek"), JsonObject), "api_key"),
                .DeepSeekBaseUrl = ReadString(TryCast(GetChildObject(plannerObject, "deepseek"), JsonObject), "base_url", DefaultDeepSeekBaseUrl),
                .DeepSeekModel = ReadString(TryCast(GetChildObject(plannerObject, "deepseek"), JsonObject), "model", DefaultDeepSeekModel),
                .OpenAiPlannerApiKey = ReadString(plannerOpenAiObject, "api_key"),
                .OpenAiPlannerBaseUrl = ReadString(plannerOpenAiObject, "base_url", DefaultOpenAiBaseUrl),
                .OpenAiPlannerModel = ReadString(plannerOpenAiObject, "model", DefaultOpenAiPlannerModel),
                .QwenPlannerApiKey = ReadString(plannerQwenObject, "api_key"),
                .QwenPlannerBaseUrl = ReadString(plannerQwenObject, "base_url", DefaultQwenBaseUrl),
                .QwenPlannerModel = ReadString(plannerQwenObject, "model", DefaultQwenPlannerModel),
                .SelectedImageProviderKey = ReadString(imageObject, "selected_provider", ResolveDefaultImageProviderKey(root, plannerObject, imageObject)),
                .OpenAiImageApiKey = ReadString(imageOpenAiObject, "api_key"),
                .OpenAiImageBaseUrl = ReadString(imageOpenAiObject, "base_url", DefaultOpenAiBaseUrl),
                .OpenAiImageModel = ReadString(imageOpenAiObject, "model", DefaultOpenAiImageModel),
                .QwenImageApiKey = ReadString(imageQwenObject, "api_key", ReadString(plannerQwenObject, "api_key")),
                .QwenImageBaseUrl = ReadString(imageQwenObject, "base_url", DefaultQwenImageEndpoint),
                .QwenImageModel = ReadString(imageQwenObject, "model", DefaultQwenImageModel),
                .GeminiApiKey = ReadString(TryCast(GetChildObject(imageObject, "gemini"), JsonObject), "api_key"),
                .GeminiImageModel = ReadString(TryCast(GetChildObject(imageObject, "gemini"), JsonObject), "model", ReadString(vertexObject, "image_model", DefaultGeminiImageModel))
            }
        End Function

        Private Shared Function ResolveDefaultPlannerProviderKey(plannerObject As JsonObject) As String
            If plannerObject Is Nothing Then
                Return "deepseek"
            End If

            Dim deepSeekApiKey = ReadString(TryCast(GetChildObject(plannerObject, "deepseek"), JsonObject), "api_key")
            If Not String.IsNullOrWhiteSpace(deepSeekApiKey) Then
                Return "deepseek"
            End If

            Dim openAiApiKey = ReadString(TryCast(GetChildObject(plannerObject, "openai"), JsonObject), "api_key")
            If Not String.IsNullOrWhiteSpace(openAiApiKey) Then
                Return "openai"
            End If

            Dim qwenApiKey = ReadString(TryCast(GetChildObject(plannerObject, "qwen"), JsonObject), "api_key")
            If Not String.IsNullOrWhiteSpace(qwenApiKey) Then
                Return "qwen"
            End If

            Return "deepseek"
        End Function

        Private Shared Function ResolveDefaultImageProviderKey(root As JsonObject,
                                                               plannerObject As JsonObject,
                                                               imageObject As JsonObject) As String
            Dim openAiApiKey = ReadString(TryCast(GetChildObject(imageObject, "openai"), JsonObject), "api_key")
            If Not String.IsNullOrWhiteSpace(openAiApiKey) Then
                Return "openai"
            End If

            Dim qwenApiKey = ReadString(TryCast(GetChildObject(imageObject, "qwen"), JsonObject), "api_key",
                                        ReadString(TryCast(GetChildObject(plannerObject, "qwen"), JsonObject), "api_key"))
            If Not String.IsNullOrWhiteSpace(qwenApiKey) Then
                Return "qwen"
            End If

            Dim geminiApiKey = ReadString(TryCast(GetChildObject(imageObject, "gemini"), JsonObject), "api_key")
            If Not String.IsNullOrWhiteSpace(geminiApiKey) OrElse HasVertexServiceAccountConfiguration(TryCast(root("vertex"), JsonObject)) Then
                Return "gemini"
            End If

            Return "gemini"
        End Function

        Private Shared Function BuildModelProviderEditors(categoryKey As String, categoryObject As JsonObject) As List(Of ModelProviderEditor)
            Dim providers As New List(Of ModelProviderEditor)()
            If categoryObject Is Nothing Then
                Return providers
            End If

            For Each propertyItem In categoryObject
                Dim providerObject = TryCast(propertyItem.Value, JsonObject)
                If providerObject Is Nothing OrElse String.Equals(propertyItem.Key, "selected_provider", StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                Dim providerEditor As New ModelProviderEditor With {
                    .CategoryKey = categoryKey,
                    .ProviderKey = propertyItem.Key,
                    .DisplayName = ReadString(providerObject, "display_name", HumanizeToken(propertyItem.Key)),
                    .Hint = ReadString(providerObject, "hint")
                }

                Dim modelOptions = ReadStringArray(providerObject, "model_options")
                For Each fieldKey In GetProviderFieldKeys(providerObject)
                    Dim fieldEditor As New ModelProviderFieldEditor With {
                        .FieldKey = fieldKey,
                        .DisplayName = GetFieldDisplayName(fieldKey),
                        .Value = ReadString(providerObject, fieldKey)
                    }

                    If String.Equals(fieldKey, "model", StringComparison.OrdinalIgnoreCase) Then
                        fieldEditor.ReplaceOptions(EnsureCurrentOption(modelOptions, fieldEditor.Value))
                    End If

                    providerEditor.Fields.Add(fieldEditor)
                Next

                providers.Add(providerEditor)
            Next

            Return providers
        End Function

        Private Shared Sub ApplyModelProviderEditors(categoryObject As JsonObject, providers As IEnumerable(Of ModelProviderEditor))
            If categoryObject Is Nothing OrElse providers Is Nothing Then
                Return
            End If

            For Each provider In providers
                If provider Is Nothing OrElse String.IsNullOrWhiteSpace(provider.ProviderKey) Then
                    Continue For
                End If

                Dim providerObject = EnsureObject(categoryObject, provider.ProviderKey)
                For Each field In provider.Fields
                    If field Is Nothing OrElse String.IsNullOrWhiteSpace(field.FieldKey) Then
                        Continue For
                    End If

                    providerObject(field.FieldKey) = If(field.Value, String.Empty)
                Next
            Next
        End Sub

        Private Shared Function GetProviderFieldKeys(providerObject As JsonObject) As IEnumerable(Of String)
            Dim keys = providerObject.
                Where(Function(item) IsEditableProviderField(item.Key, item.Value)).
                Select(Function(item) item.Key).
                ToList()

            keys.Sort(Function(left, right) CompareProviderFieldKeys(left, right))
            Return keys
        End Function

        Private Shared Function IsEditableProviderField(key As String, value As JsonNode) As Boolean
            If String.IsNullOrWhiteSpace(key) OrElse value Is Nothing Then
                Return False
            End If

            If String.Equals(key, "display_name", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(key, "model_options", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(key, "hint", StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            Return TypeOf value Is JsonValue
        End Function

        Private Shared Function CompareProviderFieldKeys(left As String, right As String) As Integer
            Dim leftRank = GetProviderFieldRank(left)
            Dim rightRank = GetProviderFieldRank(right)
            If leftRank <> rightRank Then
                Return leftRank.CompareTo(rightRank)
            End If

            Return StringComparer.OrdinalIgnoreCase.Compare(left, right)
        End Function

        Private Shared Function GetProviderFieldRank(fieldKey As String) As Integer
            Select Case fieldKey?.Trim().ToLowerInvariant()
                Case "model"
                    Return 0
                Case "base_url"
                    Return 1
                Case "api_key"
                    Return 2
                Case Else
                    Return 10
            End Select
        End Function

        Private Shared Function GetFieldDisplayName(fieldKey As String) As String
            Select Case fieldKey?.Trim().ToLowerInvariant()
                Case "model"
                    Return "Model"
                Case "base_url"
                    Return "Base URL"
                Case "api_key"
                    Return "API Key"
                Case Else
                    Return HumanizeToken(fieldKey)
            End Select
        End Function

        Private Shared Function HumanizeToken(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then
                Return String.Empty
            End If

            Dim parts = value.Split("_"c, "-"c).
                Select(Function(item) item.Trim()).
                Where(Function(item) item.Length > 0).
                Select(Function(item) HumanizeWord(item))

            Return String.Join(" ", parts)
        End Function

        Private Shared Function HumanizeWord(value As String) As String
            Select Case value.ToLowerInvariant()
                Case "api"
                    Return "API"
                Case "url"
                    Return "URL"
                Case "openai"
                    Return "OpenAI"
                Case "qwen"
                    Return "Qwen"
                Case "gemini"
                    Return "Gemini"
                Case "deepseek"
                    Return "DeepSeek"
                Case Else
                    If value.Length = 1 Then
                        Return value.ToUpperInvariant()
                    End If

                    Return Char.ToUpperInvariant(value(0)) & value.Substring(1).ToLowerInvariant()
            End Select
        End Function

        Private Shared Function ReadStringArray(parent As JsonObject, propertyName As String) As IReadOnlyList(Of String)
            Dim values = New List(Of String)()
            Dim arrayNode = TryCast(GetChildNode(parent, propertyName), JsonArray)
            If arrayNode Is Nothing Then
                Return values
            End If

            For Each item In arrayNode
                Dim textValue = item?.ToString()
                If Not String.IsNullOrWhiteSpace(textValue) Then
                    values.Add(textValue.Trim())
                End If
            Next

            Return values
        End Function

        Private Shared Function EnsureCurrentOption(values As IEnumerable(Of String), currentValue As String) As IEnumerable(Of String)
            Dim options = If(values, Enumerable.Empty(Of String)()).
                Where(Function(item) Not String.IsNullOrWhiteSpace(item)).
                ToList()

            Dim normalizedCurrent = If(currentValue, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(normalizedCurrent) AndAlso
               Not options.Any(Function(item) String.Equals(item, normalizedCurrent, StringComparison.OrdinalIgnoreCase)) Then
                options.Add(normalizedCurrent)
            End If

            Return options
        End Function

        Private Shared Function EnsureObject(parent As JsonObject, propertyName As String) As JsonObject
            Dim child = TryCast(parent(propertyName), JsonObject)
            If child Is Nothing Then
                child = New JsonObject()
                parent(propertyName) = child
            End If

            Return child
        End Function

        Private Shared Function GetChildObject(parent As JsonObject, propertyName As String) As JsonObject
            If parent Is Nothing Then
                Return Nothing
            End If

            Return TryCast(parent(propertyName), JsonObject)
        End Function

        Private Shared Function GetChildNode(parent As JsonObject, propertyName As String) As JsonNode
            If parent Is Nothing Then
                Return Nothing
            End If

            Return parent(propertyName)
        End Function

        Private Shared Function ResolveConfigPath() As String
            Dim configPath = FileLocator.FindUpwards(PreferredLocalConfigFileName)
            If String.IsNullOrWhiteSpace(configPath) Then
                configPath = FileLocator.FindUpwards(PreferredConfigFileName)
            End If
            If String.IsNullOrWhiteSpace(configPath) Then
                configPath = FileLocator.FindUpwards(LegacyLocalConfigFileName)
            End If
            If String.IsNullOrWhiteSpace(configPath) Then
                configPath = FileLocator.FindUpwards(LegacyConfigFileName)
            End If
            If String.IsNullOrWhiteSpace(configPath) Then
                Throw New FileNotFoundException($"未找到 {PreferredLocalConfigFileName}、{PreferredConfigFileName}、{LegacyLocalConfigFileName} 或 {LegacyConfigFileName}。请将配置文件放在项目目录或其上级目录。")
            End If

            Return configPath
        End Function

        Private Async Function RequestPlannerCompletionAsync(images As IEnumerable(Of ProductImageItem),
                                                             request As AnalyzeRequest,
                                                             plannerProvider As PlannerProviderSettings,
                                                             plannerModel As String,
                                                             cancellationToken As Threading.CancellationToken,
                                                             log As Action(Of String)) As Task(Of String)
            Dim endpoint = $"{plannerProvider.BaseUrl}/chat/completions"
            Dim messages = BuildPlannerMessages(images, request, plannerProvider.SupportsVision)
            log?.Invoke($"规划请求地址：{endpoint}")

            Dim requestBody As New JsonObject From {
                {"model", plannerModel},
                {"messages", messages},
                {"temperature", 0.4},
                {"response_format", New JsonObject From {
                    {"type", "json_object"}
                }}
            }

            Dim responseJson = Await PostPlannerChatCompletionAsync(endpoint, plannerProvider, requestBody, cancellationToken, log)
            Return ExtractChatCompletionText(responseJson)
        End Function

        Private Shared Function BuildPlannerMessages(images As IEnumerable(Of ProductImageItem),
                                                     request As AnalyzeRequest,
                                                     supportsVision As Boolean) As JsonArray
            Dim messages As New JsonArray()
            messages.Add(New JsonObject From {
                {"role", "system"},
                {"content", BuildAnalysisSystemPrompt(request)}
            })

            If supportsVision Then
                Dim contentParts As New JsonArray()

                For Each item In images
                    contentParts.Add(BuildPlannerVisionImagePart(item.FullPath))
                Next

                contentParts.Add(New JsonObject From {
                    {"type", "text"},
                    {"text", BuildAnalysisUserPrompt(request, images, supportsVision)}
                })

                messages.Add(New JsonObject From {
                    {"role", "user"},
                    {"content", contentParts}
                })
            Else
                messages.Add(New JsonObject From {
                    {"role", "user"},
                    {"content", BuildAnalysisUserPrompt(request, images, supportsVision)}
                })
            End If

            Return messages
        End Function

        Private Async Function PostPlannerChatCompletionAsync(endpoint As String,
                                                              plannerProvider As PlannerProviderSettings,
                                                              body As JsonObject,
                                                              cancellationToken As Threading.CancellationToken,
                                                              log As Action(Of String)) As Task(Of JsonObject)
            Using requestMessage As New HttpRequestMessage(HttpMethod.Post, endpoint)
                requestMessage.Headers.Authorization = New AuthenticationHeaderValue("Bearer", plannerProvider.ApiKey)
                requestMessage.Content = New StringContent(body.ToJsonString(JsonOptions), Encoding.UTF8, "application/json")

                Using response = Await HttpClientInstance.SendAsync(requestMessage, cancellationToken)
                    Dim content = Await response.Content.ReadAsStringAsync(cancellationToken)
                    If Not response.IsSuccessStatusCode Then
                        log?.Invoke($"规划请求失败：status={CInt(response.StatusCode)} reason={response.ReasonPhrase} body={content}")
                        Throw New HttpRequestException($"{plannerProvider.ProviderName} 请求失败：{CInt(response.StatusCode)} {response.ReasonPhrase}{Environment.NewLine}{content}")
                    End If

                    Dim node = JsonNode.Parse(content)?.AsObject()
                    If node Is Nothing Then
                        Throw New InvalidOperationException($"{plannerProvider.ProviderName} 返回了空响应。")
                    End If

                    Return node
                End Using
            End Using
        End Function

        Private Async Function PostQwenImageGenerationAsync(provider As ImageProviderSettings,
                                                            body As JsonObject,
                                                            cancellationToken As Threading.CancellationToken,
                                                            log As Action(Of String)) As Task(Of JsonObject)
            log?.Invoke($"Qwen 出图请求：endpoint={provider.BaseUrl}")

            Using requestMessage As New HttpRequestMessage(HttpMethod.Post, provider.BaseUrl)
                requestMessage.Headers.Authorization = New AuthenticationHeaderValue("Bearer", provider.ApiKey)
                requestMessage.Content = New StringContent(body.ToJsonString(JsonOptions), Encoding.UTF8, "application/json")

                Using response = Await HttpClientInstance.SendAsync(requestMessage, cancellationToken)
                    Dim content = Await response.Content.ReadAsStringAsync(cancellationToken)
                    If Not response.IsSuccessStatusCode Then
                        log?.Invoke($"Qwen 出图失败：status={CInt(response.StatusCode)} reason={response.ReasonPhrase} body={content}")
                        Throw New HttpRequestException($"Qwen 图片生成失败：{CInt(response.StatusCode)} {response.ReasonPhrase}{Environment.NewLine}{content}")
                    End If

                    Dim node = JsonNode.Parse(content)?.AsObject()
                    If node Is Nothing Then
                        Throw New InvalidOperationException("Qwen 图片生成返回了空响应。")
                    End If

                    Return node
                End Using
            End Using
        End Function

        Private Async Function PostVertexGenerateContentAsync(modelName As String,
                                                              body As JsonObject,
                                                              cancellationToken As Threading.CancellationToken,
                                                              log As Action(Of String)) As Task(Of JsonObject)
            Dim accessToken = Await GetAccessTokenAsync(cancellationToken)
            Dim endpoint = $"https://{_location}-aiplatform.googleapis.com/v1/projects/{_projectId}/locations/{_location}/publishers/google/models/{modelName}:generateContent"
            If String.Equals(_location, "global", StringComparison.OrdinalIgnoreCase) Then
                endpoint = $"https://aiplatform.googleapis.com/v1/projects/{_projectId}/locations/global/publishers/google/models/{modelName}:generateContent"
            End If

            log?.Invoke($"Vertex 出图请求：endpoint={endpoint}")

            Using requestMessage As New HttpRequestMessage(HttpMethod.Post, endpoint)
                requestMessage.Headers.Authorization = New AuthenticationHeaderValue("Bearer", accessToken)
                requestMessage.Content = New StringContent(body.ToJsonString(JsonOptions), Encoding.UTF8, "application/json")

                Using response = Await HttpClientInstance.SendAsync(requestMessage, cancellationToken)
                    Dim content = Await response.Content.ReadAsStringAsync(cancellationToken)
                    If Not response.IsSuccessStatusCode Then
                        log?.Invoke($"Vertex 出图失败：status={CInt(response.StatusCode)} reason={response.ReasonPhrase} body={content}")
                        Throw New HttpRequestException($"Vertex AI 请求失败：{CInt(response.StatusCode)} {response.ReasonPhrase}{Environment.NewLine}{content}")
                    End If

                    Dim node = JsonNode.Parse(content)?.AsObject()
                    If node Is Nothing Then
                        Throw New InvalidOperationException("Vertex AI 返回了空响应。")
                    End If

                    Return node
                End Using
            End Using
        End Function

        Private Async Function PostGeminiGenerateContentAsync(imageProvider As ImageProviderSettings,
                                                              body As JsonObject,
                                                              cancellationToken As Threading.CancellationToken,
                                                              log As Action(Of String)) As Task(Of JsonObject)
            Dim endpoint = $"{imageProvider.BaseUrl}/models/{Global.System.Uri.EscapeDataString(imageProvider.Model)}:generateContent"
            log?.Invoke($"Gemini API 出图请求：endpoint={endpoint}")

            Using requestMessage As New HttpRequestMessage(HttpMethod.Post, endpoint)
                requestMessage.Headers.Add("x-goog-api-key", imageProvider.ApiKey)
                requestMessage.Content = New StringContent(body.ToJsonString(JsonOptions), Encoding.UTF8, "application/json")

                Using response = Await HttpClientInstance.SendAsync(requestMessage, cancellationToken)
                    Dim content = Await response.Content.ReadAsStringAsync(cancellationToken)
                    If Not response.IsSuccessStatusCode Then
                        log?.Invoke($"Gemini API 出图失败：status={CInt(response.StatusCode)} reason={response.ReasonPhrase} body={content}")
                        Throw New HttpRequestException($"Gemini API 请求失败：{CInt(response.StatusCode)} {response.ReasonPhrase}{Environment.NewLine}{content}")
                    End If

                    Dim node = JsonNode.Parse(content)?.AsObject()
                    If node Is Nothing Then
                        Throw New InvalidOperationException("Gemini API 返回了空响应。")
                    End If

                    Return node
                End Using
            End Using
        End Function

        Private Async Function PostOpenAiImageEditAsync(imageProvider As ImageProviderSettings,
                                                        body As MultipartFormDataContent,
                                                        cancellationToken As Threading.CancellationToken,
                                                        log As Action(Of String)) As Task(Of JsonObject)
            Dim endpoint = $"{imageProvider.BaseUrl}/images/edits"
            log?.Invoke($"OpenAI compatible image request: endpoint={endpoint}")

            Using requestMessage As New HttpRequestMessage(HttpMethod.Post, endpoint)
                requestMessage.Headers.Authorization = New AuthenticationHeaderValue("Bearer", imageProvider.ApiKey)
                requestMessage.Content = body

                Using response = Await HttpClientInstance.SendAsync(requestMessage, cancellationToken)
                    Dim content = Await response.Content.ReadAsStringAsync(cancellationToken)
                    If Not response.IsSuccessStatusCode Then
                        log?.Invoke($"OpenAI compatible image request failed: status={CInt(response.StatusCode)} reason={response.ReasonPhrase} body={content}")
                        Throw New HttpRequestException($"OpenAI compatible image request failed: {CInt(response.StatusCode)} {response.ReasonPhrase}{Environment.NewLine}{content}")
                    End If

                    Dim node = JsonNode.Parse(content)?.AsObject()
                    If node Is Nothing Then
                        Throw New InvalidOperationException("OpenAI compatible image request returned an empty response.")
                    End If

                    Return node
                End Using
            End Using
        End Function

        Private Shared Async Function DownloadBinaryAsync(url As String,
                                                          cancellationToken As Threading.CancellationToken) As Task(Of Byte())
            Using response = Await HttpClientInstance.GetAsync(url, cancellationToken)
                response.EnsureSuccessStatusCode()
                Return Await response.Content.ReadAsByteArrayAsync(cancellationToken)
            End Using
        End Function

        Private Async Function GetAccessTokenAsync(cancellationToken As Threading.CancellationToken) As Task(Of String)
            Dim credential = GoogleCredential.FromJson(_serviceAccountJson)
            Dim scopedCredential = credential.CreateScoped(CloudPlatformScope)
            Return Await scopedCredential.UnderlyingCredential.GetAccessTokenForRequestAsync(Nothing, cancellationToken)
        End Function

        Private Sub EnsureVertexImageConfigurationAvailable(imageProvider As ImageProviderSettings)
            If imageProvider Is Nothing OrElse Not imageProvider.UsesVertex Then
                Return
            End If

            If String.IsNullOrWhiteSpace(_serviceAccountJson) Then
                Throw New InvalidOperationException($"Gemini 未配置 API Key，且 Vertex 配置文件缺少 vertex.service_account。请在系统设置中填写 Gemini API Key，或补全 {PreferredLocalConfigFileName} / {PreferredConfigFileName}。")
            End If

            If String.IsNullOrWhiteSpace(_projectId) Then
                Throw New InvalidOperationException($"Vertex 配置文件缺少 project_id。请在 {PreferredLocalConfigFileName} / {PreferredConfigFileName} 的 vertex.project_id 或 vertex.service_account.project_id 中补全。")
            End If
        End Sub

        Private Shared Function BuildAnalysisSystemPrompt(request As AnalyzeRequest) As String
            Dim schemaText = BuildAnalysisSchema().ToJsonString(JsonOptions)
            Dim useEnglish = UsesEnglishPlanning(request)
            Dim scenarioName = If(String.IsNullOrWhiteSpace(request.ScenarioName), If(useEnglish, "General Commercial Scenario", "通用商业场景"), request.ScenarioName)
            Dim scenarioDescription = If(String.IsNullOrWhiteSpace(request.ScenarioDescription), If(useEnglish, "No extra scenario description.", "无额外场景说明。"), request.ScenarioDescription)
            Dim designPlanningTemplate = If(String.IsNullOrWhiteSpace(request.ScenarioDesignPlanningTemplate),
                                            If(useEnglish, "Provide a clear, professional, and actionable design plan.", "请输出清晰、专业、可执行的设计规划。"),
                                            request.ScenarioDesignPlanningTemplate)
            Dim imagePlanningTemplate = If(String.IsNullOrWhiteSpace(request.ScenarioImagePlanningTemplate),
                                           If(useEnglish, "Provide an image plan that covers core selling points and scenario layers.", "请输出覆盖核心卖点与场景层次的图片规划。"),
                                           request.ScenarioImagePlanningTemplate)

            If useEnglish Then
                Dim englishOverlayInstruction = If(IsNoTextLanguage(request.TargetLanguage),
                                                   "If the target language is No Text, overlayTitle and overlaySubtitle must be empty strings.",
                                                   $"If the target language is not No Text, overlayTitle and overlaySubtitle must contain the on-image copy in {request.TargetLanguage}.")

                Return $"
Return a planning result that strictly follows the JSON schema.
The scenario templates below are the primary planning instructions. Follow them unless they conflict with the schema or the explicit output rules.

Scenario:
- Name: {scenarioName}
- Description: {scenarioDescription}

Design planning template:
{designPlanningTemplate}

Image planning template:
{imagePlanningTemplate}

Output rules:
1. Return JSON only. Do not output Markdown, code fences, or any extra explanation.
2. First analyze the product category, structure, materials, selling points, target audience, and use scenarios.
3. Output one unified design system and generate {request.RequestedCount} image plans.
4. Each image plan must include sequence, title, purpose, angle, scene, sellingPoint, overlayTitle, overlaySubtitle, and prompt.
5. {englishOverlayInstruction}
6. Except for overlayTitle and overlaySubtitle, every other field must be written in English.
7. The final result must be a JSON object that can be parsed directly by the application.

JSON schema:
{schemaText}
".Trim()
            End If

            Dim overlayInstruction As String
            If IsNoTextLanguage(request.TargetLanguage) Then
                overlayInstruction = "如果目标语言为无文字，overlayTitle 和 overlaySubtitle 必须返回空字符串。"
            Else
                overlayInstruction = $"如果目标语言不是无文字，overlayTitle 和 overlaySubtitle 必须填写图片上要展示的文字，且文字内容必须使用{request.TargetLanguage}。"
            End If

            Return $"
请输出一个严格符合 JSON schema 的规划结果。
下面的场景模板是主要规划依据。除非与 schema 或明确输出规则冲突，否则应优先遵循这些模板。

场景信息：
- 名称：{scenarioName}
- 描述：{scenarioDescription}

设计规划模板：
{designPlanningTemplate}

图片规划模板：
{imagePlanningTemplate}

输出规则：
1. 仅返回 JSON，不要输出 Markdown，不要输出代码块，不要输出额外解释。
2. 必须先分析产品品类、结构、材质、卖点、适用人群、使用场景。
3. 输出统一的设计规范，并生成 {request.RequestedCount} 张图片规划。
4. 每张图片规划都必须包含 sequence、title、purpose、angle、scene、sellingPoint、overlayTitle、overlaySubtitle、prompt。
5. {overlayInstruction}
6. 除 overlayTitle 和 overlaySubtitle 外，其余所有字段内容必须使用简体中文输出。
7. 返回结果必须是可直接被程序解析的 JSON object。

JSON schema:
{schemaText}
".Trim()
        End Function

        Private Shared Function BuildAnalysisUserPrompt(request As AnalyzeRequest,
                                                        images As IEnumerable(Of ProductImageItem),
                                                        supportsVision As Boolean) As String
            Dim useEnglish = UsesEnglishPlanning(request)
            Dim separator = If(useEnglish, ", ", "、")
            Dim fileHints = String.Join(separator, images.Select(Function(item) item.FileName))

            If useEnglish Then
                Dim englishImageHint = If(supportsVision,
                                          "Please analyze the attached product images directly.",
                                          "The current planning channel does not support image understanding, so rely on the requirement text and uploaded file names.")

                Return $"
{englishImageHint}
Uploaded files: {fileHints}
Project requirements: {request.RequirementText}
Business scenario: {If(String.IsNullOrWhiteSpace(request.ScenarioName), "General commercial scenario", request.ScenarioName)}
Scenario description: {If(String.IsNullOrWhiteSpace(request.ScenarioDescription), "No extra description", request.ScenarioDescription)}
Target language: {request.TargetLanguage}
Aspect ratio: {request.AspectRatio}
Requested count: {request.RequestedCount}

Return only the JSON result that matches the system instructions.
".Trim()
            End If

            Dim imageHint = If(supportsVision,
                               "请直接分析我附带的产品图片。",
                               "当前规划通道未接入图片理解，只能结合用户需求文本与上传文件名辅助判断。")

            Return $"
{imageHint}
上传文件名：{fileHints}
组图要求：{request.RequirementText}
业务场景：{If(String.IsNullOrWhiteSpace(request.ScenarioName), "通用商业场景", request.ScenarioName)}
场景说明：{If(String.IsNullOrWhiteSpace(request.ScenarioDescription), "无额外说明", request.ScenarioDescription)}
目标语言：{request.TargetLanguage}
尺寸比例：{request.AspectRatio}
生成数量：{request.RequestedCount}

请只返回符合 system 指令的 JSON 结果。
".Trim()
        End Function

        Private Shared Function BuildImagePrompt(plan As ImagePlanItem, targetLanguage As String, scenarioName As String, scenarioImageInstruction As String) As String
            Dim textInstruction = If(IsNoTextLanguage(targetLanguage),
                                     "- 画面中不要出现任何可读文字，包括标题、副标题、包装文字、参数标注和水印",
                                     $"- 输出语言：{targetLanguage}")
            Dim scenarioText = If(String.IsNullOrWhiteSpace(scenarioName), "通用商业场景", scenarioName)
            Return $"
你是电商详情图生成助手。请严格参考用户上传的产品图片，保留产品的主体结构、材质、关键部件比例和核心识别特征，不要擅自改变产品造型。
请生成一张电商详情图，满足以下要求：
- 当前场景：{scenarioText}
- 标题：{plan.Title}
- 目的：{plan.Purpose}
- 视角：{plan.Angle}
- 场景：{plan.Scene}
- 核心卖点：{plan.SellingPoint}
- 视觉文案主标题：{plan.OverlayTitle}
- 视觉文案副标题：{plan.OverlaySubtitle}
{textInstruction}

额外生成要求：
- 构图适合电商详情页，主体清晰，信息层次鲜明
- 保持高级、专业、商业摄影质感
- 如果画面适合展示文案区域，请在版面上预留易读留白
- 重点突出卖点，但不要遮挡产品主体
- 生成单张高质量成品图

补充创意提示：{plan.Prompt}
".Trim()
        End Function

        Private Shared Function UsesEnglishPlanning(request As AnalyzeRequest) As Boolean
            Dim languageCode = If(request?.PlanningLanguageCode, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(languageCode) Then
                Return False
            End If

            Return String.Equals(languageCode, "en", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function IsNoTextLanguage(value As String) As Boolean
            Dim normalized = If(value, String.Empty).Trim()
            Return String.Equals(normalized, "No Text", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(normalized, "无文字", StringComparison.Ordinal)
        End Function

        Private Shared Function ExtractQwenImageUrl(responseJson As JsonObject) As String
            Dim outputObject = TryCast(GetChildObject(responseJson, "output"), JsonObject)
            Dim choices = TryCast(GetChildNode(outputObject, "choices"), JsonArray)
            If choices Is Nothing OrElse choices.Count = 0 Then
                Return String.Empty
            End If

            Dim firstChoice = TryCast(choices(0), JsonObject)
            Dim messageObject = TryCast(GetChildObject(firstChoice, "message"), JsonObject)
            Dim content = TryCast(GetChildNode(messageObject, "content"), JsonArray)
            If content Is Nothing Then
                Return String.Empty
            End If

            For Each item In content
                Dim imageObject = TryCast(item, JsonObject)
                Dim imageValue = TryCast(GetChildNode(imageObject, "image"), JsonValue)
                Dim imageUrl = If(imageValue Is Nothing, String.Empty, imageValue.GetValue(Of String)())
                If Not String.IsNullOrWhiteSpace(imageUrl) Then
                    Return imageUrl
                End If
            Next

            Return String.Empty
        End Function

        Private Shared Async Function ExtractOpenAiImageBytesAsync(responseJson As JsonObject,
                                                                   cancellationToken As Threading.CancellationToken) As Task(Of Byte())
            Dim dataArray = TryCast(GetChildNode(responseJson, "data"), JsonArray)
            If dataArray Is Nothing OrElse dataArray.Count = 0 Then
                Return Array.Empty(Of Byte)()
            End If

            Dim firstItem = TryCast(dataArray(0), JsonObject)
            Dim imageBase64 = ReadString(firstItem, "b64_json")
            If Not String.IsNullOrWhiteSpace(imageBase64) Then
                Return Convert.FromBase64String(imageBase64)
            End If

            Dim imageUrl = ReadString(firstItem, "url")
            If Not String.IsNullOrWhiteSpace(imageUrl) Then
                Return Await DownloadBinaryAsync(imageUrl, cancellationToken)
            End If

            Return Array.Empty(Of Byte)()
        End Function

        Private Shared Function BuildDataUrl(filePath As String) As String
            Dim bytes = File.ReadAllBytes(filePath)
            Dim mimeType = GetMimeTypeFromExtension(Path.GetExtension(filePath))
            Return $"data:{mimeType};base64,{Convert.ToBase64String(bytes)}"
        End Function

        Private Shared Function MapOpenAiImageSize(aspectRatio As String, imageSize As String) As String
            Dim normalizedAspectRatio = If(aspectRatio, String.Empty).Trim()
            Dim normalizedImageSize = If(imageSize, String.Empty).Trim().ToUpperInvariant()

            If normalizedAspectRatio = "16:9" OrElse normalizedAspectRatio = "21:9" Then
                Return If(normalizedImageSize = "4K", "1536x1024", "1536x1024")
            End If

            If normalizedAspectRatio = "9:16" OrElse normalizedAspectRatio = "2:3" OrElse normalizedAspectRatio = "4:5" Then
                Return If(normalizedImageSize = "4K", "1024x1536", "1024x1536")
            End If

            Return "1024x1024"
        End Function

        Private Shared Function GetMimeTypeFromExtension(extension As String) As String
            Select Case extension?.Trim().ToLowerInvariant()
                Case ".jpg", ".jpeg"
                    Return "image/jpeg"
                Case ".bmp"
                    Return "image/bmp"
                Case ".webp"
                    Return "image/webp"
                Case ".gif"
                    Return "image/gif"
                Case ".tif", ".tiff"
                    Return "image/tiff"
                Case Else
                    Return "image/png"
            End Select
        End Function

        Private Shared Function NormalizeQwenImageSize(value As String) As String
            Dim normalized = If(value, String.Empty).Trim().ToLowerInvariant().Replace("x", "*")
            Dim parts = normalized.Split("*"c)
            If parts.Length <> 2 Then
                Throw New InvalidOperationException("Qwen 绘图尺寸格式应为 宽*高，例如 1536*1024。")
            End If

            Dim width As Integer
            Dim height As Integer
            If Not Integer.TryParse(parts(0), width) OrElse Not Integer.TryParse(parts(1), height) Then
                Throw New InvalidOperationException("Qwen 绘图尺寸必须是整数，格式示例：1536*1024。")
            End If

            If width <= 0 OrElse height <= 0 Then
                Throw New InvalidOperationException("Qwen 绘图尺寸必须为正整数。")
            End If

            Dim pixels = CLng(width) * CLng(height)
            Dim minPixels = 512L * 512L
            Dim maxPixels = 2048L * 2048L
            If pixels < minPixels OrElse pixels > maxPixels Then
                Throw New InvalidOperationException("Qwen 绘图尺寸总像素必须在 512*512 到 2048*2048 之间。")
            End If

            Return $"{width}*{height}"
        End Function

        Private Shared Function BuildAnalysisSchema() As JsonObject
            Return JsonNode.Parse(
"{
  ""type"": ""object"",
  ""properties"": {
    ""productSummary"": { ""type"": ""string"" },
    ""designTheme"": { ""type"": ""string"" },
    ""audience"": { ""type"": ""string"" },
    ""colorSystem"": { ""type"": ""string"" },
    ""typography"": { ""type"": ""string"" },
    ""visualLanguage"": { ""type"": ""string"" },
    ""photographyStyle"": { ""type"": ""string"" },
    ""layoutGuidance"": { ""type"": ""string"" },
    ""imagePlans"": {
      ""type"": ""array"",
      ""items"": {
        ""type"": ""object"",
        ""properties"": {
          ""sequence"": { ""type"": ""integer"" },
          ""title"": { ""type"": ""string"" },
          ""purpose"": { ""type"": ""string"" },
          ""angle"": { ""type"": ""string"" },
          ""scene"": { ""type"": ""string"" },
          ""sellingPoint"": { ""type"": ""string"" },
          ""overlayTitle"": { ""type"": ""string"" },
          ""overlaySubtitle"": { ""type"": ""string"" },
          ""prompt"": { ""type"": ""string"" }
        },
        ""required"": [""sequence"", ""title"", ""purpose"", ""angle"", ""scene"", ""sellingPoint"", ""overlayTitle"", ""overlaySubtitle"", ""prompt""]
      }
    }
  },
  ""required"": [""productSummary"", ""designTheme"", ""audience"", ""colorSystem"", ""typography"", ""visualLanguage"", ""photographyStyle"", ""layoutGuidance"", ""imagePlans""]
}").AsObject()
        End Function

        Private Shared Function BuildTextPart(text As String) As JsonObject
            Return New JsonObject From {
                {"text", text}
            }
        End Function

        Private Shared Function BuildImagePart(filePath As String) As JsonObject
            Dim bytes = File.ReadAllBytes(filePath)
            Dim base64 = Convert.ToBase64String(bytes)

            Return New JsonObject From {
                {"inlineData", New JsonObject From {
                    {"mimeType", GetMimeType(filePath)},
                    {"data", base64}
                }}
            }
        End Function

        Private Shared Function BuildGeminiApiImagePart(filePath As String) As JsonObject
            Dim bytes = File.ReadAllBytes(filePath)
            Dim base64 = Convert.ToBase64String(bytes)

            Return New JsonObject From {
                {"inline_data", New JsonObject From {
                    {"mime_type", GetMimeType(filePath)},
                    {"data", base64}
                }}
            }
        End Function

        Private Shared Function BuildPlannerVisionImagePart(filePath As String) As JsonObject
            Dim base64 = Convert.ToBase64String(File.ReadAllBytes(filePath))
            Dim mimeType = GetMimeType(filePath)

            Return New JsonObject From {
                {"type", "image_url"},
                {"image_url", New JsonObject From {
                    {"url", $"data:{mimeType};base64,{base64}"}
                }}
            }
        End Function

        Private Shared Function GetMimeType(filePath As String) As String
            Select Case Path.GetExtension(filePath).ToLowerInvariant()
                Case ".jpg", ".jpeg"
                    Return "image/jpeg"
                Case ".webp"
                    Return "image/webp"
                Case ".bmp"
                    Return "image/bmp"
                Case ".tif", ".tiff"
                    Return "image/tiff"
                Case ".heic"
                    Return "image/heic"
                Case Else
                    Return "image/png"
            End Select
        End Function

        Private Shared Function ExtractChatCompletionText(response As JsonObject) As String
            Dim messageContentNode = response("choices")?.AsArray()?(0)?("message")?("content")
            If messageContentNode Is Nothing Then
                Return String.Empty
            End If

            If TypeOf messageContentNode Is JsonValue Then
                Dim textValue = messageContentNode.GetValue(Of String)()
                Return ExtractJsonPayload(textValue)
            End If

            Dim contentArray = TryCast(messageContentNode, JsonArray)
            If contentArray Is Nothing Then
                Return String.Empty
            End If

            Dim builder As New StringBuilder()
            For Each item In contentArray
                Dim text = item?("text")?.GetValue(Of String)()
                If Not String.IsNullOrWhiteSpace(text) Then
                    builder.AppendLine(text)
                End If
            Next

            Return ExtractJsonPayload(builder.ToString())
        End Function

        Private Shared Function ExtractJsonPayload(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then
                Return String.Empty
            End If

            Dim trimmed = value.Trim()
            If trimmed.StartsWith("```", StringComparison.Ordinal) Then
                trimmed = trimmed.Trim("`"c).Trim()
                If trimmed.StartsWith("json", StringComparison.OrdinalIgnoreCase) Then
                    trimmed = trimmed.Substring(4).Trim()
                End If
            End If

            Dim startIndex = trimmed.IndexOf("{"c)
            Dim endIndex = trimmed.LastIndexOf("}"c)
            If startIndex >= 0 AndAlso endIndex > startIndex Then
                Return trimmed.Substring(startIndex, endIndex - startIndex + 1)
            End If

            Return trimmed
        End Function

        Private Shared Function ExtractImageData(response As JsonObject) As Byte()
            Dim parts = response("candidates")?.AsArray()?(0)?("content")?("parts")?.AsArray()
            If parts Is Nothing Then
                Return Array.Empty(Of Byte)()
            End If

            For Each part In parts
                Dim data = ReadJsonString(part, "inlineData", "data")
                If String.IsNullOrWhiteSpace(data) Then
                    data = ReadJsonString(part, "inline_data", "data")
                End If

                If Not String.IsNullOrWhiteSpace(data) Then
                    Return Convert.FromBase64String(data)
                End If
            Next

            Return Array.Empty(Of Byte)()
        End Function

        Private Shared Function ReadJsonString(parent As JsonNode, childName As String, valueName As String) As String
            Dim parentObject = TryCast(parent, JsonObject)
            If parentObject Is Nothing Then
                Return String.Empty
            End If

            Dim child = TryCast(parentObject(childName), JsonObject)
            If child Is Nothing OrElse child(valueName) Is Nothing Then
                Return String.Empty
            End If

            Dim value = child(valueName)?.GetValue(Of String)()
            Return If(value, String.Empty)
        End Function

        Private Shared Function ParseDesignPlan(jsonText As String) As DesignPlan
            Dim root = JsonNode.Parse(jsonText)?.AsObject()
            If root Is Nothing Then
                Throw New InvalidOperationException("无法解析规划模型返回的规划 JSON。")
            End If

            Dim plan As New DesignPlan With {
                .ProductSummary = root("productSummary")?.GetValue(Of String)(),
                .DesignTheme = root("designTheme")?.GetValue(Of String)(),
                .Audience = root("audience")?.GetValue(Of String)(),
                .ColorSystem = root("colorSystem")?.GetValue(Of String)(),
                .Typography = root("typography")?.GetValue(Of String)(),
                .VisualLanguage = root("visualLanguage")?.GetValue(Of String)(),
                .PhotographyStyle = root("photographyStyle")?.GetValue(Of String)(),
                .LayoutGuidance = root("layoutGuidance")?.GetValue(Of String)()
            }

            Dim imagePlans = root("imagePlans")?.AsArray()
            If imagePlans IsNot Nothing Then
                For Each item In imagePlans
                    If item Is Nothing Then
                        Continue For
                    End If

                    Dim planItem = item.AsObject()
                    Dim sequenceNode = planItem("sequence")
                    Dim sequence As Integer = 0
                    If sequenceNode IsNot Nothing Then
                        sequence = sequenceNode.GetValue(Of Integer)()
                    End If

                    plan.ImagePlans.Add(New ImagePlanItem With {
                        .Sequence = sequence,
                        .Title = planItem("title")?.GetValue(Of String)(),
                        .Purpose = planItem("purpose")?.GetValue(Of String)(),
                        .Angle = planItem("angle")?.GetValue(Of String)(),
                        .Scene = planItem("scene")?.GetValue(Of String)(),
                        .SellingPoint = planItem("sellingPoint")?.GetValue(Of String)(),
                        .OverlayTitle = planItem("overlayTitle")?.GetValue(Of String)(),
                        .OverlaySubtitle = planItem("overlaySubtitle")?.GetValue(Of String)(),
                        .Prompt = planItem("prompt")?.GetValue(Of String)()
                    })
                Next
            End If

            Return plan
        End Function

        Private Shared Function ReadString(obj As JsonObject, propertyName As String, Optional defaultValue As String = "") As String
            If obj Is Nothing OrElse Not obj.ContainsKey(propertyName) OrElse obj(propertyName) Is Nothing Then
                Return defaultValue
            End If

            Dim value = obj(propertyName)?.GetValue(Of String)()
            If String.IsNullOrWhiteSpace(value) Then
                Return defaultValue
            End If

            Return value.Trim()
        End Function

        Private Shared Function NormalizeBaseUrl(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then
                Return String.Empty
            End If

            Return value.Trim().TrimEnd("/"c)
        End Function

        Private Shared Function NormalizeEndpoint(value As String, fallback As String) As String
            If String.IsNullOrWhiteSpace(value) Then
                Return fallback
            End If

            Return value.Trim()
        End Function

        Private Shared Sub EnsureImageGenerationModelSupported(modelName As String)
            If String.IsNullOrWhiteSpace(modelName) Then
                Throw New InvalidOperationException($"当前未配置图片生成模型。请在 {PreferredConfigFileName} 的 vertex.image_model 中填写支持图片输出的模型。")
            End If

            Dim normalized = modelName.Trim().ToLowerInvariant()
            Dim looksLikeTextOnlyGemini = normalized.Contains("gemini") AndAlso
                                          Not normalized.Contains("image") AndAlso
                                          Not normalized.Contains("imagen")

            If looksLikeTextOnlyGemini Then
                Throw New InvalidOperationException($"当前图片生成模型 {modelName} 不支持图片输出。请改用支持出图的模型，例如 gemini-2.5-flash-image。")
            End If
        End Sub

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

        Private Shared Function SanitizeFileName(value As String) As String
            Dim invalid = Path.GetInvalidFileNameChars()
            Dim builder As New StringBuilder()

            For Each ch In value
                If invalid.Contains(ch) Then
                    builder.Append("_"c)
                Else
                    builder.Append(ch)
                End If
            Next

            Return builder.ToString().Trim().Trim("."c)
        End Function

        Private Shared Function ClonePlannerProvider(provider As PlannerProviderSettings) As PlannerProviderSettings
            If provider Is Nothing Then
                Return Nothing
            End If

            Return New PlannerProviderSettings With {
                .ProviderKey = provider.ProviderKey,
                .ProviderName = provider.ProviderName,
                .ApiKey = provider.ApiKey,
                .BaseUrl = provider.BaseUrl,
                .Model = provider.Model,
                .SupportsVision = provider.SupportsVision
            }
        End Function

        Private Shared Function CloneImageProvider(provider As ImageProviderSettings) As ImageProviderSettings
            If provider Is Nothing Then
                Return Nothing
            End If

            Return New ImageProviderSettings With {
                .ProviderKey = provider.ProviderKey,
                .ProviderName = provider.ProviderName,
                .ApiKey = provider.ApiKey,
                .BaseUrl = provider.BaseUrl,
                .Model = provider.Model,
                .RequiresFlexibleSize = provider.RequiresFlexibleSize,
                .UsesVertex = provider.UsesVertex
            }
        End Function

        Private Shared Function HasVertexServiceAccountConfiguration(vertexObject As JsonObject) As Boolean
            If vertexObject Is Nothing Then
                Return False
            End If

            Dim serviceAccountObject = TryCast(vertexObject("service_account"), JsonObject)
            If serviceAccountObject Is Nothing Then
                Return False
            End If

            Dim projectId = ReadString(vertexObject, "project_id", ReadString(serviceAccountObject, "project_id"))
            Return Not String.IsNullOrWhiteSpace(projectId)
        End Function

        Private NotInheritable Class PlannerProviderSettings
            Public Property ProviderKey As String
            Public Property ProviderName As String
            Public Property ApiKey As String
            Public Property BaseUrl As String
            Public Property Model As String
            Public Property SupportsVision As Boolean
        End Class

        Private NotInheritable Class ImageProviderSettings
            Public Property ProviderKey As String
            Public Property ProviderName As String
            Public Property ApiKey As String
            Public Property BaseUrl As String
            Public Property Model As String
            Public Property RequiresFlexibleSize As Boolean
            Public Property UsesVertex As Boolean
        End Class
    End Class
End Namespace

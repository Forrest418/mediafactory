Imports System.Collections.Generic
Imports MediaFactory.Infrastructure

Namespace Models
    Public Class ModelProviderConfiguration
        Inherits ObservableObject

        Private _selectedPlannerProviderKey As String = "deepseek"
        Private _selectedImageProviderKey As String = "gemini"
        Private _deepSeekApiKey As String = String.Empty
        Private _deepSeekBaseUrl As String = "https://api.deepseek.com"
        Private _deepSeekModel As String = "deepseek-chat"
        Private _openAiPlannerApiKey As String = String.Empty
        Private _openAiPlannerBaseUrl As String = "https://api.openai.com/v1"
        Private _openAiPlannerModel As String = "gpt-4.1-mini"
        Private _qwenPlannerApiKey As String = String.Empty
        Private _qwenPlannerBaseUrl As String = "https://dashscope.aliyuncs.com/compatible-mode/v1"
        Private _qwenPlannerModel As String = "qwen3.5-vl-plus"
        Private _openAiImageApiKey As String = String.Empty
        Private _openAiImageBaseUrl As String = "https://api.openai.com/v1"
        Private _openAiImageModel As String = "gpt-image-1.5"
        Private _qwenImageApiKey As String = String.Empty
        Private _qwenImageBaseUrl As String = "https://dashscope.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation"
        Private _qwenImageModel As String = "qwen-image-2.0-pro"
        Private _geminiApiKey As String = String.Empty
        Private _geminiImageModel As String = "gemini-3.1-flash-image-preview"

        Public Property SelectedPlannerProviderKey As String
            Get
                Return _selectedPlannerProviderKey
            End Get
            Set(value As String)
                SetProperty(_selectedPlannerProviderKey, NormalizeKey(value, "deepseek"))
            End Set
        End Property

        Public Property SelectedImageProviderKey As String
            Get
                Return _selectedImageProviderKey
            End Get
            Set(value As String)
                SetProperty(_selectedImageProviderKey, NormalizeKey(value, "gemini"))
            End Set
        End Property

        Public Property DeepSeekApiKey As String
            Get
                Return _deepSeekApiKey
            End Get
            Set(value As String)
                SetProperty(_deepSeekApiKey, NormalizeText(value))
            End Set
        End Property

        Public Property DeepSeekBaseUrl As String
            Get
                Return _deepSeekBaseUrl
            End Get
            Set(value As String)
                SetProperty(_deepSeekBaseUrl, NormalizeText(value))
            End Set
        End Property

        Public Property DeepSeekModel As String
            Get
                Return _deepSeekModel
            End Get
            Set(value As String)
                SetProperty(_deepSeekModel, NormalizeText(value))
            End Set
        End Property

        Public Property OpenAiPlannerApiKey As String
            Get
                Return _openAiPlannerApiKey
            End Get
            Set(value As String)
                SetProperty(_openAiPlannerApiKey, NormalizeText(value))
            End Set
        End Property

        Public Property OpenAiPlannerBaseUrl As String
            Get
                Return _openAiPlannerBaseUrl
            End Get
            Set(value As String)
                SetProperty(_openAiPlannerBaseUrl, NormalizeText(value))
            End Set
        End Property

        Public Property OpenAiPlannerModel As String
            Get
                Return _openAiPlannerModel
            End Get
            Set(value As String)
                SetProperty(_openAiPlannerModel, NormalizeText(value))
            End Set
        End Property

        Public Property QwenPlannerApiKey As String
            Get
                Return _qwenPlannerApiKey
            End Get
            Set(value As String)
                SetProperty(_qwenPlannerApiKey, NormalizeText(value))
            End Set
        End Property

        Public Property QwenPlannerBaseUrl As String
            Get
                Return _qwenPlannerBaseUrl
            End Get
            Set(value As String)
                SetProperty(_qwenPlannerBaseUrl, NormalizeText(value))
            End Set
        End Property

        Public Property QwenPlannerModel As String
            Get
                Return _qwenPlannerModel
            End Get
            Set(value As String)
                SetProperty(_qwenPlannerModel, NormalizeText(value))
            End Set
        End Property

        Public Property OpenAiImageApiKey As String
            Get
                Return _openAiImageApiKey
            End Get
            Set(value As String)
                SetProperty(_openAiImageApiKey, NormalizeText(value))
            End Set
        End Property

        Public Property OpenAiImageBaseUrl As String
            Get
                Return _openAiImageBaseUrl
            End Get
            Set(value As String)
                SetProperty(_openAiImageBaseUrl, NormalizeText(value))
            End Set
        End Property

        Public Property OpenAiImageModel As String
            Get
                Return _openAiImageModel
            End Get
            Set(value As String)
                SetProperty(_openAiImageModel, NormalizeText(value))
            End Set
        End Property

        Public Property QwenImageApiKey As String
            Get
                Return _qwenImageApiKey
            End Get
            Set(value As String)
                SetProperty(_qwenImageApiKey, NormalizeText(value))
            End Set
        End Property

        Public Property QwenImageBaseUrl As String
            Get
                Return _qwenImageBaseUrl
            End Get
            Set(value As String)
                SetProperty(_qwenImageBaseUrl, NormalizeText(value))
            End Set
        End Property

        Public Property QwenImageModel As String
            Get
                Return _qwenImageModel
            End Get
            Set(value As String)
                SetProperty(_qwenImageModel, NormalizeText(value))
            End Set
        End Property

        Public Property GeminiApiKey As String
            Get
                Return _geminiApiKey
            End Get
            Set(value As String)
                SetProperty(_geminiApiKey, NormalizeText(value))
            End Set
        End Property

        Public Property GeminiImageModel As String
            Get
                Return _geminiImageModel
            End Get
            Set(value As String)
                SetProperty(_geminiImageModel, NormalizeText(value))
            End Set
        End Property

        Public Function Clone() As ModelProviderConfiguration
            Return CType(MemberwiseClone(), ModelProviderConfiguration)
        End Function

        Private Shared Function NormalizeText(value As String) As String
            Return If(value, String.Empty).Trim()
        End Function

        Private Shared Function NormalizeKey(value As String, fallback As String) As String
            Dim normalized = NormalizeText(value).ToLowerInvariant()
            Return If(String.IsNullOrWhiteSpace(normalized), fallback, normalized)
        End Function
    End Class

    Public Class ProviderOptionItem
        Public Property Key As String = String.Empty
        Public Property DisplayName As String = String.Empty
    End Class
End Namespace

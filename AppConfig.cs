using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SmartMailAssistant1
{
    public class AppConfig
    {
        public ApiSettings ApiSettings { get; set; }
        public UiSettings UiSettings { get; set; }
        public WindowSettings WindowSettings { get; set; }
        public EmailSettings EmailSettings { get; set; }
        public LanguageSettings LanguageSettings { get; set; }
        public ReplySettings ReplySettings { get; set; }
        public Messages Messages { get; set; }
        public Prompts Prompts { get; set; }
        public RibbonSettings RibbonSettings { get; set; } 
        public OutlookSettings OutlookSettings { get; set; } 
        public ClassificationSettings ClassificationSettings { get; set; }
    }

    public class ApiSettings
    {
        public string ApiKey { get; set; }
        public string ApiEndpoint { get; set; }
        public string ModelName { get; set; }
        public int TimeoutSeconds { get; set; }
        public string MessageRole { get; set; }
        public string StreamPrefix { get; set; }
        public string StreamEndMarker { get; set; }
        public string ResponseContentPath { get; set; }
        public Dictionary<string, string> RequestHeaders { get; set; }
        public Dictionary<string, object> RequestParameters { get; set; }

        [JsonIgnore]
        public string ApiUrl
        {
            get => ApiEndpoint;
            set => ApiEndpoint = value;
        }
    }

    public class UiSettings
    {
        public FontSettings Fonts { get; set; }
        public ColorSettings Colors { get; set; }
        public ButtonSettings ButtonSettings { get; set; }
        public LoadingSettings LoadingSettings { get; set; }
        public StreamingSettings StreamingSettings { get; set; }
    }

    public class FontSettings
    {
        public string Default { get; set; }
        public string Arabic { get; set; }
        public string Hebrew { get; set; }
        public string Persian { get; set; }
        public string Urdu { get; set; }
        public string Pashto { get; set; }
        public string DefaultRtl { get; set; }
    }

    public class ColorSettings
    {
        public string DarkBackground { get; set; }
        public string TextColor { get; set; }
        public string TextBoxBackground { get; set; }
        public string RetryButtonColor { get; set; }
        public string SendButtonColor { get; set; }
        public string GenerateAnywayButtonColor { get; set; }
        public string CloseButtonColor { get; set; }
    }

    public class ButtonSettings
    {
        public int PanelHeight { get; set; }
        public int ButtonWidth { get; set; }
        public int LargeButtonWidth { get; set; }
        public int ButtonSpacing { get; set; }
        public int CloseButtonSpacing { get; set; }
        public int ForceReplySendButtonWidth { get; set; }
        public Dictionary<string, string> CustomActions { get; set; }
    }

    public class LoadingSettings
    {
        public int AnimationInterval { get; set; }
        public List<string> LoadingStates { get; set; }
    }
    public class OutlookSettings
    {
        public string MapiNamespace { get; set; }
        public DefaultFoldersSettings DefaultFolders { get; set; }
        public ItemRestrictions ItemRestrictions { get; set; }

        public string DefaultFallbackFolder { get; set; } // Config-driven only
        public ViewSettings ViewSettings { get; set; }
    }

    public class DefaultFoldersSettings
    {
        public string PrimaryFolder { get; set; }
        public string FallbackFolder { get; set; }
        public List<string> SupportedFolderTypes { get; set; }
    }

    public class ItemRestrictions
    {
        public string UnreadFilter { get; set; }
        public string ReadFilter { get; set; }
        public string TodayFilter { get; set; }
        public string ThisWeekFilter { get; set; }
        public List<string> CustomFilters { get; set; }
    }

    public class ViewSettings
    {
        public bool ResetViewAfterClassification { get; set; }
        public bool RefreshViewAfterChanges { get; set; }
        public int ViewUpdateDelayMs { get; set; }
    }

    public class StreamingSettings
    {
        public int CharacterDelay { get; set; }
    }

    public class WindowSettings
    {
        public WindowSize MainWindow { get; set; }
        public ProgressWindowSettings ProgressWindow { get; set; } 
    }

    public class WindowSize
    {
        public int Width { get; set; }
        public int Height { get; set; }
    }

    
    public class ProgressWindowSettings
    {
        public int Width { get; set; }
        public int Height { get; set; }
        public int MinHeight { get; set; }
        public int HeightDivisor { get; set; }
        public string FormBorderStyle { get; set; }
        public string StartPosition { get; set; }
        public bool MaximizeBox { get; set; }
        public bool MinimizeBox { get; set; }
    }

    public class EmailSettings
    {
        public List<string> MessageSeparators { get; set; }
        public List<string> SignaturePrefixes { get; set; }
        public DefaultFolders DefaultFolders { get; set; } 
        public ContentLimits ContentLimits { get; set; } 
        public EmptyContentDefaults EmptyContentDefaults { get; set; } 
    }

    
    public class DefaultFolders
    {
        public string Primary { get; set; }
        public string Fallback { get; set; }
    }

    public class ContentLimits
    {
        public int MaxOriginalBodyPreview { get; set; }
        public int MaxProcessedBodyPreview { get; set; }
        public int MinimalContentLength { get; set; }
        public int FallbackBodyLength { get; set; }
    }

    public class EmptyContentDefaults
    {
        public string NoSubject { get; set; }
        public string NoSender { get; set; }
        public string NoBody { get; set; }
    }

    public class LanguageSettings
    {
        public string DefaultTranslationLanguage { get; set; }
        public List<string> RightToLeftLanguages { get; set; }
        public string FallbackLanguage { get; set; }
    }

    public class ReplySettings
    {
        public List<string> NoReplyTriggers { get; set; }
    }

    public class Messages
    {
        public string Loading { get; set; }
        public string Retry { get; set; }
        public string Close { get; set; }
        public string SendReply { get; set; }
        public string GenerateAnyway { get; set; }
        public string Complete { get; set; }
        public string Error { get; set; }
        public string ClickRetry { get; set; }
        public string NoReplyNeeded { get; set; }
        public string DefaultReplyText { get; set; }
        public string ApiError { get; set; }
        public string GeneratingReply { get; set; }
        public string ErrorCreatingReply { get; set; }
        public string NoEmailSelected { get; set; }
        public string NoEmailSelectedTitle { get; set; }
        public string EmptyEmailBody { get; set; }
        public string EmptyEmailTitle { get; set; }
        public string GeneratingSummary { get; set; }
        public string ErrorGeneratingSummary { get; set; }
        public string Translating { get; set; }
        public string ErrorTranslating { get; set; }
        public string AnalyzingEmail { get; set; }
        public string ErrorAnalyzing { get; set; }
        public string InitializationError { get; set; }
        public string CheckConfigFile { get; set; }
        public string ConfigurationError { get; set; }
        public string MissingApiKey { get; set; }
    }

    public class Prompts
    {
        public string Summarize { get; set; }
        public string Translate { get; set; }
        public string Reply { get; set; }
        public string ForceReply { get; set; }
    }

    public class ClassificationSettings
    {
        public string Prompt { get; set; }
        public Dictionary<string, ClassificationConfig> Classifications { get; set; }
        public RateLimitSettings RateLimiting { get; set; }
        public PreProcessingRules PreProcessingRules { get; set; } = new PreProcessingRules();
        public AiClassificationSettings AiClassification { get; set; } = new AiClassificationSettings();
        public EmailProcessingSettings EmailProcessing { get; set; }
        public ClassificationMessages Messages { get; set; }
        public DebugSettings Debug { get; set; }
        public RetrySettings Retry { get; set; }
        public ContentSettings Content { get; set; }
        public ApiResponseSettings ApiResponse { get; set; }
        public ProgressSettings Progress { get; set; }
        public NormalizationSettings Normalization { get; set; }
        public SecurityProtocolSettings SecurityProtocol { get; set; } 
    }

    public class ClassificationConfig
    {
        public string CategoryPrefix { get; set; }
        public string Importance { get; set; }
        public string FlagIcon { get; set; }
        public string FlagRequest { get; set; }
        public string SubjectPrefix { get; set; }
        public string CategoryColor { get; set; }
    }

    public class RateLimitSettings
    {

        public int DefaultRequestTimeoutSeconds { get; set; } // Config-driven only

        public int DefaultMinDelayBetweenRequestsMs { get; set; } // Config-driven only
        public int MaxConsecutiveFailures { get; set; }
        public int RequestsPerMinute { get; set; }
        public int BaseCooldownSeconds { get; set; }
        public int MaxCooldownMinutes { get; set; }
        public int MinDelayBetweenRequestsMs { get; set; }
        public int RequestTimeoutSeconds { get; set; }
    }

    public class PreProcessingRules
    {
        public bool EnableRuleBasedClassification { get; set; } = false;
        public List<ClassificationRule> Rules { get; set; } = new List<ClassificationRule>();
    }

    public class ClassificationRule
    {
        public string Name { get; set; }
        public string Classification { get; set; }
        public int Priority { get; set; } = 100;
        public RuleConditions Conditions { get; set; } = new RuleConditions();
    }

    public class RuleConditions
    {
        public string MatchType { get; set; } = "ANY";
        public List<string> SubjectKeywords { get; set; } = new List<string>();
        public List<string> SenderDomains { get; set; } = new List<string>();
        public List<string> SenderAddresses { get; set; } = new List<string>();
        public List<string> BodyKeywords { get; set; } = new List<string>();
    }

    public class AiClassificationSettings
    {
        public bool EnableAiClassification { get; set; } = true;
        public bool UseAiForUnmatched { get; set; } = true;
        public DynamicPrompts DynamicPrompts { get; set; } = new DynamicPrompts();
    }

    public class DynamicPrompts
    {
        public bool IncludeAvailableCategories { get; set; } = false;
        public string PromptTemplate { get; set; }
        public Dictionary<string, string> CategoryDescriptions { get; set; } = new Dictionary<string, string>();
    }

    public class EmailProcessingSettings
    {
        public int DefaultMaxBodyLength { get; set; } // Config-driven only
        public List<string> ContentSeparators { get; set; }
        public int MaxBodyLength { get; set; }
    }

    public class ClassificationMessages
    {
        public string NoUnreadEmails { get; set; }
        public string ConfirmClassification { get; set; }
        public string CompletionSummary { get; set; }
        public string HowToSeeClassifications { get; set; }
        public string ProcessingStatus { get; set; }
        public string CooldownStatus { get; set; }
        public string RetryStatus { get; set; }
        public string WaitStatus { get; set; }
    }

    public class DebugSettings
    {

        public int DefaultContentPreviewLength { get; set; } // Config-driven only
        public bool EnableLogging { get; set; }
        public int ContentPreviewLength { get; set; }
        public string TruncationIndicator { get; set; }

        public string DefaultTruncationIndicator { get; set; } // Config-driven only
    }

    public class RetrySettings
    {
        public int MaxAttempts { get; set; }
        public List<string> DefaultFallbackOrder { get; set; }
        public BackoffMultipliers BackoffMultipliers { get; set; } 
    }

    // NEW: Backoff multipliers
    public class BackoffMultipliers
    {
        public int TimeoutRetryBase { get; set; }
        public int ErrorRetryBase { get; set; }
        public int ExponentialBase { get; set; }
        public int ExtendedCooldownFactor { get; set; } 



    }

    public class ContentSettings
    {
        public List<string> AiTagPatterns { get; set; }
        public string FallbackContentTemplate { get; set; }
        public List<string> ClassificationKeywords { get; set; }
        public List<string> DefaultSeparators { get; set; } 

        public int EmptyContentFallbackLength { get; set; }
        public string MinimalContentFromSubjectTemplate { get; set; } 
        
    }

    public class ApiResponseSettings
    {
        public List<string> FallbackContentPaths { get; set; }
        public List<string> NoContentIndicators { get; set; }
    }

    public class ProgressSettings
    {
        public string TimeFormatString { get; set; }
        public string StatusTemplate { get; set; }
        public string CooldownTemplate { get; set; }
        public string PercentageTemplate { get; set; }
        public int UpdateIntervalMs { get; set; } 
        public string ProgressBarStyle { get; set; } 
    }

    
    public class NormalizationSettings
    {
        public Dictionary<string, string> ImportanceSynonyms { get; set; }
        public Dictionary<string, string> FlagIconSynonyms { get; set; }
        public Dictionary<string, string> CategoryColorSynonyms { get; set; }

        public string DefaultImportance { get; set; } // Config-driven only
        public string DefaultCategoryColor { get; set; } // Config-driven only
        public string DefaultFlagIcon { get; set; } // Config-driven only
    }

    public class RibbonSettings
    {
        public RibbonButtonSettings Ribbon1 { get; set; }
        public RibbonButtonSettings Ribbon2 { get; set; }
    }

    public class RibbonButtonSettings
    {
        public string Button1Text { get; set; }
        public string Button1Tooltip { get; set; }
        public string Button2Text { get; set; }
        public string Button2Tooltip { get; set; }
        public string Button3Text { get; set; }
        public string Button3Tooltip { get; set; }
        public string Button5Text { get; set; }
        public string Button5Tooltip { get; set; }
        
        public string Button4Text { get; set; }        
        public string Button4Tooltip { get; set; }
    }

    // NEW: Security protocol settings
    public class SecurityProtocolSettings
    {
        public string PrimaryProtocol { get; set; }
        public string FallbackProtocol { get; set; }
        public bool TryTls13 { get; set; }
    }
}
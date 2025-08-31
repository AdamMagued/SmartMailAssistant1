using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SmartMailAssistant1
{
    public static class ConfigManager
    {
        private static readonly string ConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
        private static AppConfig _config;
        private static bool _configValid = false;
        private static readonly List<string> _validationErrors = new List<string>();

        public static AppConfig Config
        {
            get
            {
                if (!_configValid) LoadConfig();
                return _configValid ? _config : null;
            }
        }

        public static bool IsConfigValid
        {
            get
            {
                if (!_configValid) LoadConfig();
                return _configValid;
            }
        }

        public static IReadOnlyList<string> ValidationErrors => _validationErrors;

        public static void LoadConfig()
        {
            _configValid = false;
            _validationErrors.Clear();

            try
            {
                if (!File.Exists(ConfigPath))
                {
                    _validationErrors.Add($"Config file not found at: {ConfigPath}");
                    return;
                }

                string json = File.ReadAllText(ConfigPath);
                if (string.IsNullOrWhiteSpace(json))
                {
                    _validationErrors.Add("Config file is empty");
                    return;
                }

                _config = JsonConvert.DeserializeObject<AppConfig>(json);
                if (_config == null)
                {
                    _validationErrors.Add("Failed to deserialize config file");
                    return;
                }

                _configValid = ValidateConfig(_config);
            }
            catch (JsonException ex)
            {
                _validationErrors.Add($"JSON parsing error: {ex.Message}");
            }
            catch (Exception ex)
            {
                _validationErrors.Add($"Error loading config: {ex.Message}");
            }
        }

        public static bool SaveConfig()
        {
            try
            {
                string json = JsonConvert.SerializeObject(_config, Formatting.Indented);
                File.WriteAllText(ConfigPath, json);
                _configValid = true;
                return true;
            }
            catch (Exception ex)
            {
                _validationErrors.Add($"Error saving config: {ex.Message}");
                return false;
            }
        }

        private static bool ValidateConfig(AppConfig config)
        {
            bool isValid = true;

            isValid &= ValidateNotNull(config.ApiSettings, nameof(config.ApiSettings));
            isValid &= ValidateNotNull(config.OutlookSettings, nameof(config.OutlookSettings));
            isValid &= ValidateNotNull(config.UiSettings, nameof(config.UiSettings));
            isValid &= ValidateNotNull(config.WindowSettings, nameof(config.WindowSettings));
            isValid &= ValidateNotNull(config.EmailSettings, nameof(config.EmailSettings));
            isValid &= ValidateNotNull(config.LanguageSettings, nameof(config.LanguageSettings));
            isValid &= ValidateNotNull(config.ReplySettings, nameof(config.ReplySettings));
            isValid &= ValidateNotNull(config.Messages, nameof(config.Messages));
            isValid &= ValidateNotNull(config.Prompts, nameof(config.Prompts));
            isValid &= ValidateNotNull(config.ClassificationSettings, nameof(config.ClassificationSettings));

            if (!isValid) return false;

            isValid &= ValidateApiSettings(config.ApiSettings);
            isValid &= ValidateOutlookSettings(config.OutlookSettings);
            isValid &= ValidateUiSettings(config.UiSettings);
            isValid &= ValidateWindowSettings(config.WindowSettings);
            isValid &= ValidateEmailSettings(config.EmailSettings);
            isValid &= ValidateLanguageSettings(config.LanguageSettings);
            isValid &= ValidateReplySettings(config.ReplySettings);
            isValid &= ValidateMessages(config.Messages);
            isValid &= ValidatePrompts(config.Prompts);
            isValid &= ValidateClassificationSettings(config.ClassificationSettings);
            isValid &= ValidateRibbonSettings(config.RibbonSettings);

            return isValid;
        }

        private static bool ValidateNotNull(object obj, string name)
        {
            if (obj == null)
            {
                _validationErrors.Add($"{name} is missing.");
                return false;
            }
            return true;
        }

        private static bool ValidateRibbonSettings(RibbonSettings ribbon)
        {
            if (ribbon == null)
            {
                _validationErrors.Add("RibbonSettings is missing.");
                return false;
            }

            bool valid = true;
            valid &= ValidateNotNull(ribbon.Ribbon1, nameof(ribbon.Ribbon1));
            valid &= ValidateNotNull(ribbon.Ribbon2, nameof(ribbon.Ribbon2));
            return valid;
        }


        private static bool ValidateNotEmpty(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                _validationErrors.Add($"{name} cannot be empty.");
                return false;
            }
            return true;
        }

        private static bool ValidateOutlookSettings(OutlookSettings outlook)
        {
            bool valid = true;
            valid &= ValidateNotEmpty(outlook.MapiNamespace, nameof(outlook.MapiNamespace));
            valid &= ValidateNotNull(outlook.DefaultFolders, nameof(outlook.DefaultFolders));
            valid &= ValidateNotNull(outlook.ItemRestrictions, nameof(outlook.ItemRestrictions));
            valid &= ValidateNotNull(outlook.ViewSettings, nameof(outlook.ViewSettings));

            if (outlook.DefaultFolders != null)
            {
                valid &= ValidateNotEmpty(outlook.DefaultFolders.PrimaryFolder, nameof(outlook.DefaultFolders.PrimaryFolder));
                valid &= ValidateNotEmpty(outlook.DefaultFolders.FallbackFolder, nameof(outlook.DefaultFolders.FallbackFolder));
                valid &= ValidateNotNull(outlook.DefaultFolders.SupportedFolderTypes, nameof(outlook.DefaultFolders.SupportedFolderTypes));
            }

            return valid;
        }


        private static bool ValidateApiSettings(ApiSettings api)
        {
            bool valid = true;
            valid &= ValidateNotEmpty(api.ApiKey, nameof(api.ApiKey));
            if (string.IsNullOrWhiteSpace(api.ApiEndpoint) && string.IsNullOrWhiteSpace(api.ApiUrl))
            {
                _validationErrors.Add("ApiSettings.ApiEndpoint/ApiUrl cannot be empty.");
                valid = false;
            }
            valid &= ValidateNotEmpty(api.ModelName, nameof(api.ModelName));
            return valid;
        }

        private static bool ValidateUiSettings(UiSettings ui)
        {
            bool valid = true;
            valid &= ValidateNotNull(ui.Fonts, nameof(ui.Fonts));
            valid &= ValidateNotNull(ui.Colors, nameof(ui.Colors));
            valid &= ValidateNotNull(ui.ButtonSettings, nameof(ui.ButtonSettings));
            valid &= ValidateNotNull(ui.LoadingSettings, nameof(ui.LoadingSettings));
            valid &= ValidateNotNull(ui.StreamingSettings, nameof(ui.StreamingSettings));

            if (ui.Fonts != null)
            {
                valid &= ValidateNotEmpty(ui.Fonts.Default, nameof(ui.Fonts.Default));
                valid &= ValidateNotEmpty(ui.Fonts.DefaultRtl, nameof(ui.Fonts.DefaultRtl)); // 🔧 ADD THIS LINE
            }
            if (ui.ButtonSettings != null)
            {
                if (ui.ButtonSettings.ForceReplySendButtonWidth <= 0)
                {
                    _validationErrors.Add("UiSettings.ButtonSettings.ForceReplySendButtonWidth must be greater than 0.");
                    valid = false;
                }
            }

            if (ui.Colors != null)
                valid &= ValidateNotEmpty(ui.Colors.TextColor, nameof(ui.Colors.TextColor));

            return valid;
        }

        private static bool ValidateWindowSettings(WindowSettings window)
        {
            bool valid = true;
            valid &= ValidateNotNull(window.MainWindow, nameof(window.MainWindow));
            valid &= ValidateNotNull(window.ProgressWindow, nameof(window.ProgressWindow));
            return valid;
        }

        private static bool ValidateEmailSettings(EmailSettings email)
        {
            bool valid = true;
            valid &= ValidateNotNull(email.MessageSeparators, nameof(email.MessageSeparators));
            valid &= ValidateNotNull(email.SignaturePrefixes, nameof(email.SignaturePrefixes));
            valid &= ValidateNotNull(email.DefaultFolders, nameof(email.DefaultFolders));
            valid &= ValidateNotNull(email.ContentLimits, nameof(email.ContentLimits));
            valid &= ValidateNotNull(email.EmptyContentDefaults, nameof(email.EmptyContentDefaults));
            return valid;
        }

        private static bool ValidateLanguageSettings(LanguageSettings lang)
        {
            bool valid = true;
            valid &= ValidateNotEmpty(lang.DefaultTranslationLanguage, nameof(lang.DefaultTranslationLanguage));
            valid &= ValidateNotEmpty(lang.FallbackLanguage, nameof(lang.FallbackLanguage)); // 🔧 ADD THIS LINE
            valid &= ValidateNotNull(lang.RightToLeftLanguages, nameof(lang.RightToLeftLanguages));
            return valid;
        }

        private static bool ValidateReplySettings(ReplySettings reply)
        {
            return ValidateNotNull(reply.NoReplyTriggers, nameof(reply.NoReplyTriggers));
        }

        private static bool ValidateMessages(Messages messages)
        {
            bool valid = true;
            foreach (var prop in typeof(Messages).GetProperties())
            {
                if (prop.GetValue(messages) is string value && string.IsNullOrWhiteSpace(value))
                {
                    _validationErrors.Add($"Messages.{prop.Name} cannot be empty.");
                    valid = false;
                }
            }
            return valid;
        }

        private static bool ValidatePrompts(Prompts prompts)
        {
            bool valid = true;
            foreach (var prop in typeof(Prompts).GetProperties())
            {
                if (prop.GetValue(prompts) is string value && string.IsNullOrWhiteSpace(value))
                {
                    _validationErrors.Add($"Prompts.{prop.Name} cannot be empty.");
                    valid = false;
                }
            }
            return valid;
        }

        private static bool ValidateClassificationSettings(ClassificationSettings cs)
        {
            bool valid = true;
            valid &= ValidateNotEmpty(cs.Prompt, nameof(cs.Prompt));
            valid &= ValidateNotNull(cs.Classifications, nameof(cs.Classifications));
            valid &= ValidateNotNull(cs.RateLimiting, nameof(cs.RateLimiting));
            valid &= ValidateNotNull(cs.PreProcessingRules, nameof(cs.PreProcessingRules));
            valid &= ValidateNotNull(cs.AiClassification, nameof(cs.AiClassification));
            valid &= ValidateNotNull(cs.EmailProcessing, nameof(cs.EmailProcessing));
            valid &= ValidateNotNull(cs.Messages, nameof(cs.Messages));
            valid &= ValidateNotNull(cs.Debug, nameof(cs.Debug));
            valid &= ValidateNotNull(cs.Retry, nameof(cs.Retry));
            valid &= ValidateNotNull(cs.Content, nameof(cs.Content));
            valid &= ValidateNotNull(cs.ApiResponse, nameof(cs.ApiResponse));
            valid &= ValidateNotNull(cs.Progress, nameof(cs.Progress));
            valid &= ValidateNotNull(cs.Normalization, nameof(cs.Normalization));
            valid &= ValidateNotNull(cs.SecurityProtocol, nameof(cs.SecurityProtocol));

            // Just validate that at least one category is defined
            if (cs.Classifications != null && cs.Classifications.Count == 0)
            {
                _validationErrors.Add("ClassificationSettings.Classifications must define at least one category.");
                valid = false;
            }

            // REMOVE THE DUPLICATE VALIDATION BLOCKS BELOW!
            // You only need each validation block ONCE

            if (cs.EmailProcessing != null)
            {
                valid &= ValidatePositiveNumber(cs.EmailProcessing.DefaultMaxBodyLength, "EmailProcessing.DefaultMaxBodyLength");
            }

            if (cs.Debug != null)
            {
                valid &= ValidatePositiveNumber(cs.Debug.DefaultContentPreviewLength, "Debug.DefaultContentPreviewLength");
                valid &= ValidateNotEmpty(cs.Debug.DefaultTruncationIndicator, "Debug.DefaultTruncationIndicator");
            }

            if (cs.Content != null)
            {
                valid &= ValidatePositiveNumber(cs.Content.EmptyContentFallbackLength, "Content.EmptyContentFallbackLength");
                valid &= ValidateNotNull(cs.Content.DefaultSeparators, "Content.DefaultSeparators");
                valid &= ValidateNotEmpty(cs.Content.MinimalContentFromSubjectTemplate, "Content.MinimalContentFromSubjectTemplate");
            }

            if (cs.RateLimiting != null)
            {
                valid &= ValidatePositiveNumber(cs.RateLimiting.DefaultMinDelayBetweenRequestsMs, "RateLimiting.DefaultMinDelayBetweenRequestsMs");
                valid &= ValidatePositiveNumber(cs.RateLimiting.DefaultRequestTimeoutSeconds, "RateLimiting.DefaultRequestTimeoutSeconds");
            }
            if (cs.Retry?.BackoffMultipliers != null)
            {
                valid &= ValidatePositiveNumber(cs.Retry.BackoffMultipliers.TimeoutRetryBase, "Retry.BackoffMultipliers.TimeoutRetryBase");
                valid &= ValidatePositiveNumber(cs.Retry.BackoffMultipliers.ErrorRetryBase, "Retry.BackoffMultipliers.ErrorRetryBase");
                valid &= ValidatePositiveNumber(cs.Retry.BackoffMultipliers.ExponentialBase, "Retry.BackoffMultipliers.ExponentialBase");
                valid &= ValidatePositiveNumber(cs.Retry.BackoffMultipliers.ExtendedCooldownFactor, "Retry.BackoffMultipliers.ExtendedCooldownFactor"); // NEW
            }
           
            if (cs.Retry?.BackoffMultipliers != null)
            {
                valid &= ValidatePositiveNumber(cs.Retry.BackoffMultipliers.TimeoutRetryBase, "Retry.BackoffMultipliers.TimeoutRetryBase");
                valid &= ValidatePositiveNumber(cs.Retry.BackoffMultipliers.ErrorRetryBase, "Retry.BackoffMultipliers.ErrorRetryBase");
                valid &= ValidatePositiveNumber(cs.Retry.BackoffMultipliers.ExponentialBase, "Retry.BackoffMultipliers.ExponentialBase");
            }

            if (cs.Normalization != null)
            {
                valid &= ValidateNotEmpty(cs.Normalization.DefaultImportance, "Normalization.DefaultImportance");
                valid &= ValidateNotEmpty(cs.Normalization.DefaultCategoryColor, "Normalization.DefaultCategoryColor");
                valid &= ValidateNotEmpty(cs.Normalization.DefaultFlagIcon, "Normalization.DefaultFlagIcon");
            }

            return valid;
        }

        private static bool ValidatePositiveNumber(int value, string name)
        {
            if (value <= 0)
            {
                _validationErrors.Add($"{name} must be greater than 0.");
                return false;
            }
            return true;
        }

        // Helpers
        public static Font ParseFont(string fontString)
        {
            if (string.IsNullOrWhiteSpace(fontString))
                throw new Exception("Font string is empty.");

            var parts = fontString.Split(',');
            if (parts.Length < 2) throw new Exception("Invalid font format");

            string name = parts[0].Trim();
            if (!float.TryParse(parts[1].Trim(), out float size))
                throw new Exception("Invalid font size");

            FontStyle style = FontStyle.Regular;
            if (parts.Length > 2 && Enum.TryParse(parts[2].Trim(), out FontStyle parsed))
                style = parsed;

            return new Font(name, size, style);
        }

        public static Color ParseColor(string colorString)
        {
            if (string.IsNullOrWhiteSpace(colorString))
                throw new Exception("Color string is empty.");

            var parts = colorString.Split(',');
            if (parts.Length != 3) throw new Exception("Color must have 3 components (R,G,B)");

            int r = int.Parse(parts[0].Trim());
            int g = int.Parse(parts[1].Trim());
            int b = int.Parse(parts[2].Trim());

            return Color.FromArgb(r, g, b);
        }
    }
}
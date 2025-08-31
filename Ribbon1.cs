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
    public partial class Ribbon1
    {
        private Font defaultFont;
        private Font arabicFont;
        private Font hebrewFont;
        private Font persianFont;
        private Font urduFont;
        private Font pashtoFont;
        private Font defaultRtlFont;
        private Color darkBackground;
        private Color textColor;
        private AppConfig config;
        private bool _initialized = false;

        // Email classification state (from Ribbon2)
        

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                LoadConfiguration();

                if (config == null)
                {
                    string errorMessage = GetConfigValue("Messages.ConfigurationFailedToLoad", "Configuration failed to load") + ":\n\n" +
                                        string.Join("\n", ConfigManager.ValidationErrors) +
                                        "\n\n" + GetConfigValue("Messages.FixConfigAndRestart", "Please fix config.txt and restart Outlook.");
                    MessageBox.Show(errorMessage, GetConfigValue("Messages.ConfigurationError", "Configuration Error"),
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    DisableAllButtons();
                    return;
                }

                InitializeUISettings();
                ConfigureSecurityProtocol();
                ValidateApiConfiguration();
                ValidateClassificationConfiguration(); // Added for email classification

                _initialized = true;
                ApplyRibbonConfiguration();

            }
            catch (Exception ex)
            {
                MessageBox.Show($"{GetConfigValue("Messages.InitializationError", "Initialization error")}: {ex.Message}\n\n{GetConfigValue("Messages.CheckConfigFile", "Please check your config file and restart Outlook.")}",
                    GetConfigValue("Messages.ConfigurationError", "Configuration Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                DisableAllButtons();
            }
        }

        private string GetConfigValue(string path, string defaultValue)
        {
            try
            {
                if (config == null) return defaultValue;

                var parts = path.Split('.');
                object current = config;

                foreach (var part in parts)
                {
                    if (current == null) return defaultValue;
                    var prop = current.GetType().GetProperty(part);
                    if (prop == null) return defaultValue;
                    current = prop.GetValue(current);
                }

                return current?.ToString() ?? defaultValue;
            }
            catch
            {
                return defaultValue;
            }
        }

        private void DisableAllButtons()
        {
            try
            {
                var buttonFields = this.GetType().GetFields(System.Reflection.BindingFlags.NonPublic |
                                                           System.Reflection.BindingFlags.Instance)
                                      .Where(f => f.FieldType.Name.Contains(GetConfigValue("UiSettings.ButtonSettings.ButtonTypeFilter", "RibbonButton")))
                                      .ToList();

                foreach (var field in buttonFields)
                {
                    var button = field.GetValue(this) as Microsoft.Office.Tools.Ribbon.RibbonButton;
                    if (button != null)
                    {
                        button.Enabled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"{GetConfigValue("Messages.ErrorDisablingButtons", "Error disabling buttons")}: {ex.Message}");
            }
        }

        private void LoadConfiguration()
        {
            config = ConfigManager.Config;
            if (config == null)
            {
                throw new Exception(GetConfigValue("Messages.ConfigFileValidationFailed", "Configuration file validation failed"));
            }
        }

        private void InitializeUISettings()
        {
            defaultFont = ConfigManager.ParseFont(config.UiSettings.Fonts.Default);
            arabicFont = ConfigManager.ParseFont(config.UiSettings.Fonts.Arabic);
            hebrewFont = ConfigManager.ParseFont(config.UiSettings.Fonts.Hebrew);
            persianFont = ConfigManager.ParseFont(config.UiSettings.Fonts.Persian);
            urduFont = ConfigManager.ParseFont(config.UiSettings.Fonts.Urdu);
            pashtoFont = ConfigManager.ParseFont(config.UiSettings.Fonts.Pashto);
            defaultRtlFont = ConfigManager.ParseFont(config.UiSettings.Fonts.DefaultRtl);
            darkBackground = ConfigManager.ParseColor(config.UiSettings.Colors.DarkBackground);
            textColor = ConfigManager.ParseColor(config.UiSettings.Colors.TextColor);
        }

        private void ConfigureSecurityProtocol()
        {
            try
            {
                var securityConfig = config.ClassificationSettings?.SecurityProtocol;
                if (securityConfig != null)
                {
                    // Use configured primary protocol
                    if (Enum.TryParse<SecurityProtocolType>(securityConfig.PrimaryProtocol, out var primaryProtocol))
                    {
                        ServicePointManager.SecurityProtocol = primaryProtocol;
                    }
                    else
                    {
                        // Fallback to configured fallback protocol
                        if (Enum.TryParse<SecurityProtocolType>(securityConfig.FallbackProtocol, out var fallbackProtocol))
                        {
                            ServicePointManager.SecurityProtocol = fallbackProtocol;
                        }
                    }

                    // Try TLS 1.3 if configured
                    if (securityConfig.TryTls13 && Enum.IsDefined(typeof(SecurityProtocolType), GetConfigValue("ClassificationSettings.SecurityProtocol.Tls13Name", "Tls13")))
                    {
                        ServicePointManager.SecurityProtocol |= (SecurityProtocolType)Enum.Parse(typeof(SecurityProtocolType), GetConfigValue("ClassificationSettings.SecurityProtocol.Tls13Name", "Tls13"));
                    }
                }
                else
                {
                    // Use hardcoded fallback only if no config exists
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                }
            }
            catch (Exception)
            {
                // Final fallback using config or default
                var fallbackProtocol = GetConfigValue("ClassificationSettings.SecurityProtocol.FallbackProtocol", "Tls12");
                if (Enum.TryParse<SecurityProtocolType>(fallbackProtocol, out var protocol))
                {
                    ServicePointManager.SecurityProtocol = protocol;
                }
                else
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                }
            }
        }

        private void ValidateApiConfiguration()
        {
            if (config == null) return;
            if (config.ApiSettings == null) throw new Exception(GetConfigValue("Messages.ApiSettingsMissing", "API Settings section is missing in configuration"));
            if (string.IsNullOrWhiteSpace(config.ApiSettings.ApiKey)) throw new Exception(config.Messages.MissingApiKey);
            config.ApiSettings.ApiKey = config.ApiSettings.ApiKey.Trim();
            if (string.IsNullOrWhiteSpace(config.ApiSettings.ApiUrl)) throw new Exception(GetConfigValue("Messages.ApiUrlMissing", "API URL is missing in configuration"));
            if (string.IsNullOrWhiteSpace(config.ApiSettings.ModelName)) throw new Exception(GetConfigValue("Messages.ModelNameMissing", "Model name is missing in configuration"));

            if (config.UiSettings.LoadingSettings.AnimationInterval <= 0)
                throw new Exception(GetConfigValue("Messages.AnimationIntervalError", "LoadingSettings.AnimationInterval must be greater than 0"));
            if (config.UiSettings.StreamingSettings.CharacterDelay <= 0)
                throw new Exception(GetConfigValue("Messages.CharacterDelayError", "StreamingSettings.CharacterDelay must be greater than 0"));

            // Validate required UI elements
            if (config.UiSettings.LoadingSettings.LoadingStates == null || !config.UiSettings.LoadingSettings.LoadingStates.Any())
                throw new Exception(GetConfigValue("Messages.LoadingStatesError", "LoadingSettings.LoadingStates must contain at least one loading state"));
            if (string.IsNullOrWhiteSpace(config.LanguageSettings.FallbackLanguage))
                throw new Exception(GetConfigValue("Messages.FallbackLanguageError", "LanguageSettings.FallbackLanguage is required"));
            if (string.IsNullOrWhiteSpace(config.UiSettings.Fonts.DefaultRtl))
                throw new Exception(GetConfigValue("Messages.DefaultRtlFontError", "UiSettings.Fonts.DefaultRtl is required"));
        }

        // Added for email classification functionality
        private void ValidateClassificationConfiguration()
        {
            // Only validate if classification settings exist (optional feature)
            if (config.ClassificationSettings == null) return;

            if (string.IsNullOrWhiteSpace(config.ClassificationSettings.Prompt))
                throw new Exception("ClassificationSettings.Prompt is missing in configuration");

            if (config.ClassificationSettings.Classifications == null ||
                config.ClassificationSettings.Classifications.Count == 0)
                throw new Exception("ClassificationSettings.Classifications must define at least one class.");

            if (config.ClassificationSettings.EmailProcessing == null)
                throw new Exception("ClassificationSettings.EmailProcessing is missing.");
            if (config.ClassificationSettings.RateLimiting == null)
                throw new Exception("ClassificationSettings.RateLimiting is missing.");
            if (config.ClassificationSettings.Messages == null)
                throw new Exception("ClassificationSettings.Messages is missing.");
        }

        private string ExtractMostRecentMessage(string emailBody)
        {
            if (string.IsNullOrWhiteSpace(emailBody)) return emailBody;
            int firstSeparatorPos = -1;
            foreach (var separator in config.EmailSettings.MessageSeparators)
            {
                int pos = emailBody.IndexOf(separator, StringComparison.OrdinalIgnoreCase);
                if (pos > 0 && (firstSeparatorPos == -1 || pos < firstSeparatorPos))
                {
                    firstSeparatorPos = pos;
                }
            }
            return firstSeparatorPos > 0 ? emailBody.Substring(0, firstSeparatorPos).Trim() : emailBody.Trim();
        }

        private string CleanMessage(string message)
        {
            if (string.IsNullOrWhiteSpace(message)) return message;
            var lines = message.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None)
                            .TakeWhile(line => !config.EmailSettings.SignaturePrefixes.Any(prefix =>
                                line.Trim().StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
                            .ToArray();
            return string.Join(GetConfigValue("EmailSettings.LineJoinSeparator", "\n"), lines).Trim();
        }

        private bool IsRightToLeftLanguage(string language)
        {
            return config.LanguageSettings.RightToLeftLanguages.Any(rtlLang =>
                language.IndexOf(rtlLang, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private Font GetFontForLanguage(string language)
        {
            if (string.IsNullOrEmpty(language)) return defaultFont;

            if (IsRightToLeftLanguage(language))
            {
                foreach (var rtlLang in config.LanguageSettings.RightToLeftLanguages)
                {
                    if (language.IndexOf(rtlLang, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        switch (rtlLang.ToLowerInvariant())
                        {
                            case "arabic": return arabicFont;
                            case "hebrew": return hebrewFont;
                            case "persian": return persianFont;
                            case "urdu": return urduFont;
                            case "pashto": return pashtoFont;
                        }
                    }
                }
                return defaultRtlFont;
            }
            return defaultFont;
        }

        private List<string> GetLoadingStatesForLanguage(string language)
        {
            return config.UiSettings.LoadingSettings.LoadingStates;
        }

        private int GetConfigInt(string dottedPath, int defaultValue)
        {
            try
            {
                if (config == null || string.IsNullOrWhiteSpace(dottedPath)) return defaultValue;
                object current = config;
                var parts = dottedPath.Split('.');
                foreach (var p in parts)
                {
                    if (current == null) return defaultValue;
                    var prop = current.GetType().GetProperty(p);
                    if (prop == null) return defaultValue;
                    current = prop.GetValue(current);
                }
                if (current == null) return defaultValue;
                if (current is int i) return i;
                if (current is long l) return Convert.ToInt32(l);
                if (current is short s) return Convert.ToInt32(s);
                if (current is string str && int.TryParse(str, out int parsed)) return parsed;
                if (current is decimal dec) return Convert.ToInt32(dec);
                if (current is double dbl) return Convert.ToInt32(dbl);
                return defaultValue;
            }
            catch { return defaultValue; }
        }

        private void ApplyRibbonConfiguration()
        {
            if (config?.RibbonSettings?.Ribbon1 == null) return;

            var ribbonConfig = config.RibbonSettings.Ribbon1;

            // Apply button 1 settings (Summarize)
            if (!string.IsNullOrWhiteSpace(ribbonConfig.Button1Text))
                SummaryButton.Label = ribbonConfig.Button1Text;
            if (!string.IsNullOrWhiteSpace(ribbonConfig.Button1Tooltip))
                SummaryButton.ScreenTip = ribbonConfig.Button1Tooltip;

            // Apply button 2 settings (Translate)
            if (!string.IsNullOrWhiteSpace(ribbonConfig.Button2Text))
                TranslateButton.Label = ribbonConfig.Button2Text;
            if (!string.IsNullOrWhiteSpace(ribbonConfig.Button2Tooltip))
                TranslateButton.ScreenTip = ribbonConfig.Button2Tooltip;

            // Apply button 3 settings (Smart Reply)
            if (!string.IsNullOrWhiteSpace(ribbonConfig.Button3Text))
                SuggestReplyButton.Label = ribbonConfig.Button3Text;
            if (!string.IsNullOrWhiteSpace(ribbonConfig.Button3Tooltip))
                SuggestReplyButton.ScreenTip = ribbonConfig.Button3Tooltip;

            // Apply button 4 settings (Classify Emails) - merged from Ribbon2
            if (!string.IsNullOrWhiteSpace(ribbonConfig.Button4Text))
                ClassifyButton.Label = ribbonConfig.Button4Text;
            else if (config.RibbonSettings?.Ribbon2?.Button5Text != null) // Fallback to old Ribbon2 config
                ClassifyButton.Label = config.RibbonSettings.Ribbon2.Button5Text;

            if (!string.IsNullOrWhiteSpace(ribbonConfig.Button4Tooltip))
                ClassifyButton.ScreenTip = ribbonConfig.Button4Tooltip;
            else if (config.RibbonSettings?.Ribbon2?.Button5Tooltip != null) // Fallback to old Ribbon2 config
                ClassifyButton.ScreenTip = config.RibbonSettings.Ribbon2.Button5Tooltip;
        }

        private async Task StreamApiResponse(string prompt, string windowTitle, string targetLanguage = null, bool isReplyMode = false, Outlook.MailItem originalMail = null)
        {
            if (config == null)
            {
                MessageBox.Show(GetConfigValue("Messages.ConfigurationNotLoaded", "Configuration not loaded"),
                    GetConfigValue("Messages.Error", "Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string fullResponse = "";
            string displayResponse = "";
            int replyCode = GetConfigInt("ReplySettings.DefaultReplyCode", -1);

            bool isRtl = !string.IsNullOrEmpty(targetLanguage) && IsRightToLeftLanguage(targetLanguage);
            Font currentFont = !string.IsNullOrEmpty(targetLanguage) ? GetFontForLanguage(targetLanguage) : defaultFont;
            List<string> loadingStates = GetLoadingStatesForLanguage(targetLanguage);

            var streamingForm = new Form
            {
                Width = config.WindowSettings.MainWindow.Width,
                Height = config.WindowSettings.MainWindow.Height,
                Text = windowTitle,
                StartPosition = (FormStartPosition)Enum.Parse(typeof(FormStartPosition),
                    GetConfigValue("WindowSettings.MainWindow.StartPosition", "CenterScreen")),
                Font = currentFont,
                BackColor = darkBackground,
                ForeColor = textColor,
                FormBorderStyle = (FormBorderStyle)Enum.Parse(typeof(FormBorderStyle),
                    GetConfigValue("WindowSettings.MainWindow.FormBorderStyle", "FixedDialog")),
                MaximizeBox = GetConfigValue("WindowSettings.MainWindow.MaximizeBox", "false").ToLower() == "true",
                MinimizeBox = GetConfigValue("WindowSettings.MainWindow.MinimizeBox", "false").ToLower() == "true",
                RightToLeftLayout = isRtl
            };

            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = darkBackground,
                RightToLeft = isRtl ? RightToLeft.Yes : RightToLeft.No
            };

            var loadingLabel = new Label
            {
                Text = loadingStates.FirstOrDefault() ?? config.Messages.Loading,
                Font = currentFont,
                ForeColor = ConfigManager.ParseColor(GetConfigValue("UiSettings.Colors.LoadingTextColor", "211,211,211")),
                BackColor = Color.Transparent,
                AutoSize = false,
                TextAlign = (ContentAlignment)Enum.Parse(typeof(ContentAlignment),
                    GetConfigValue("UiSettings.LoadingSettings.TextAlign", "MiddleCenter")),
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.No
            };

            var textBox = new RichTextBox
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = currentFont,
                BackColor = ConfigManager.ParseColor(config.UiSettings.Colors.TextBoxBackground),
                ForeColor = textColor,
                BorderStyle = (BorderStyle)Enum.Parse(typeof(BorderStyle),
                    GetConfigValue("UiSettings.TextBox.BorderStyle", "None")),
                WordWrap = GetConfigValue("UiSettings.TextBox.WordWrap", "true").ToLower() == "true",
                ScrollBars = (RichTextBoxScrollBars)Enum.Parse(typeof(RichTextBoxScrollBars),
                    GetConfigValue("UiSettings.TextBox.ScrollBars", "Vertical")),
                Visible = false,
                RightToLeft = isRtl ? RightToLeft.Yes : RightToLeft.No
            };

            var bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = config.UiSettings.ButtonSettings.PanelHeight,
                BackColor = darkBackground,
                Padding = new Padding(
                    GetConfigInt("UiSettings.ButtonSettings.PanelPadding.Left", 10),
                    GetConfigInt("UiSettings.ButtonSettings.PanelPadding.Top", 5),
                    GetConfigInt("UiSettings.ButtonSettings.PanelPadding.Right", 10),
                    GetConfigInt("UiSettings.ButtonSettings.PanelPadding.Bottom", 5))
            };

            var retryButton = new Button()
            {
                Text = config.Messages.Retry,
                Font = currentFont,
                Width = config.UiSettings.ButtonSettings.ButtonWidth,
                Height = GetConfigInt("UiSettings.ButtonSettings.ButtonHeight", 35),
                Visible = false,
                BackColor = ConfigManager.ParseColor(config.UiSettings.Colors.RetryButtonColor),
                ForeColor = textColor,
                FlatStyle = (FlatStyle)Enum.Parse(typeof(FlatStyle),
                    GetConfigValue("UiSettings.ButtonSettings.FlatStyle", "Flat")),
                RightToLeft = isRtl ? RightToLeft.Yes : RightToLeft.No
            };
            retryButton.FlatAppearance.BorderSize = GetConfigInt("UiSettings.ButtonSettings.BorderSize", 0);

            var sendReplyButton = new Button()
            {
                Text = config.Messages.SendReply,
                Font = currentFont,
                Width = GetConfigInt("UiSettings.ButtonSettings.ForceReplySendButtonWidth",
                         config.UiSettings.ButtonSettings.ButtonWidth), // fallback
                Height = GetConfigInt("UiSettings.ButtonSettings.ButtonHeight", 35),
                Visible = false,
                BackColor = ConfigManager.ParseColor(config.UiSettings.Colors.SendButtonColor),
                ForeColor = textColor,
                FlatStyle = (FlatStyle)Enum.Parse(typeof(FlatStyle),
        GetConfigValue("UiSettings.ButtonSettings.FlatStyle", "Flat"))
            };

            sendReplyButton.FlatAppearance.BorderSize = GetConfigInt("UiSettings.ButtonSettings.BorderSize", 0);

            var generateAnywayButton = new Button()
            {
                Text = config.Messages.GenerateAnyway,
                Font = currentFont,
                Width = config.UiSettings.ButtonSettings.LargeButtonWidth,
                Height = GetConfigInt("UiSettings.ButtonSettings.ButtonHeight", 35),
                Visible = false,
                BackColor = ConfigManager.ParseColor(config.UiSettings.Colors.GenerateAnywayButtonColor),
                ForeColor = textColor,
                FlatStyle = (FlatStyle)Enum.Parse(typeof(FlatStyle),
                    GetConfigValue("UiSettings.ButtonSettings.FlatStyle", "Flat"))
            };
            generateAnywayButton.FlatAppearance.BorderSize = GetConfigInt("UiSettings.ButtonSettings.BorderSize", 0);

            var closeButton = new Button()
            {
                Text = config.Messages.Close,
                Font = currentFont,
                Width = config.UiSettings.ButtonSettings.ButtonWidth,
                Height = GetConfigInt("UiSettings.ButtonSettings.ButtonHeight", 35),
                Visible = false,
                BackColor = ConfigManager.ParseColor(config.UiSettings.Colors.CloseButtonColor),
                ForeColor = textColor,
                FlatStyle = (FlatStyle)Enum.Parse(typeof(FlatStyle),
                    GetConfigValue("UiSettings.ButtonSettings.FlatStyle", "Flat")),
                RightToLeft = isRtl ? RightToLeft.Yes : RightToLeft.No
            };
            closeButton.FlatAppearance.BorderSize = GetConfigInt("UiSettings.ButtonSettings.BorderSize", 0);
            closeButton.Click += (s, ev) => streamingForm.Close();

            int buttonSpacing = GetConfigInt("UiSettings.ButtonSettings.ButtonSpacing", 10);
            int closeButtonSpacing = GetConfigInt("UiSettings.ButtonSettings.CloseButtonSpacing", 20);

            retryButton.Margin = new Padding(0, 0, buttonSpacing, 0);
            generateAnywayButton.Margin = new Padding(0, 0, buttonSpacing, 0);
            sendReplyButton.Margin = new Padding(0, 0, 0, 0);

            var leftButtonFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                FlowDirection = (FlowDirection)Enum.Parse(typeof(FlowDirection),
                    GetConfigValue("UiSettings.ButtonSettings.FlowDirection", "LeftToRight")),
                AutoSize = true,
                WrapContents = GetConfigValue("UiSettings.ButtonSettings.WrapContents", "false").ToLower() == "true",
                BackColor = darkBackground,
                Padding = new Padding(0, GetConfigInt("UiSettings.ButtonSettings.FlowPanelPadding", 10), 0, 0)
            };

            leftButtonFlow.Controls.Add(retryButton);
            leftButtonFlow.Controls.Add(generateAnywayButton);
            leftButtonFlow.Controls.Add(sendReplyButton);

            var rightPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = closeButton.Width + closeButtonSpacing + GetConfigInt("UiSettings.ButtonSettings.RightPanelExtraWidth", 10),
                BackColor = darkBackground
            };
            closeButton.Location = new Point(closeButtonSpacing / 2, (bottomPanel.Height - closeButton.Height) / 2);
            rightPanel.Controls.Add(closeButton);

            bottomPanel.Controls.Add(leftButtonFlow);
            bottomPanel.Controls.Add(rightPanel);

            var dotTimer = new System.Windows.Forms.Timer
            {
                Interval = config.UiSettings.LoadingSettings.AnimationInterval
            };
            int dotCount = 0;

            dotTimer.Tick += (s, e) =>
            {
                if (loadingStates.Count > 0)
                {
                    dotCount = (dotCount + 1) % loadingStates.Count;
                    loadingLabel.Text = loadingStates[dotCount];
                }
            };

            void SafeInvoke(Control control, Action action)
            {
                try
                {
                    if (control.InvokeRequired)
                        control.Invoke(action);
                    else
                        action();
                }
                catch (ObjectDisposedException) { }
                catch (InvalidOperationException) { }
            }

            string ParseReplyCodeAndGetDisplayText(string response)
            {
                if (string.IsNullOrWhiteSpace(response))
                {
                    replyCode = GetConfigInt("ReplySettings.NoContentReplyCode", 0);
                    return config.Messages.NoReplyNeeded;
                }

                string noReplyPrefix = GetConfigValue("ReplySettings.NoReplyPrefix", "0:");
                string replyPrefix = GetConfigValue("ReplySettings.ReplyPrefix", "1:");

                if (response.StartsWith(noReplyPrefix))
                {
                    replyCode = GetConfigInt("ReplySettings.NoReplyCode", 0);
                    return config.Messages.NoReplyNeeded;
                }
                else if (response.StartsWith(replyPrefix))
                {
                    replyCode = GetConfigInt("ReplySettings.ReplyCode", 1);
                    return response.Length > replyPrefix.Length ? response.Substring(replyPrefix.Length).Trim() : config.Messages.DefaultReplyText;
                }

                if (config.ReplySettings.NoReplyTriggers.Any(trigger =>
                    response.IndexOf(trigger, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    replyCode = GetConfigInt("ReplySettings.NoReplyCode", 0);
                    return config.Messages.NoReplyNeeded;
                }

                replyCode = GetConfigInt("ReplySettings.ReplyCode", 1);
                return response;
            }

            string ExtractContentFromResponse(dynamic data, string path)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(path))
                        return data?.ToString() ?? string.Empty;

                    var pathParts = path.Split(new[] { '[', ']', '.' }, StringSplitOptions.RemoveEmptyEntries);
                    object current = data;

                    foreach (var part in pathParts)
                    {
                        if (current == null) break;

                        if (current is Newtonsoft.Json.Linq.JObject jObject)
                        {
                            current = jObject[part];
                        }
                        else if (current is Newtonsoft.Json.Linq.JArray jArray && int.TryParse(part, out int index))
                        {
                            current = index < jArray.Count ? jArray[index] : null;
                        }
                        else
                        {
                            current = null;
                        }
                    }

                    return current?.ToString() ?? string.Empty;
                }
                catch
                {
                    return string.Empty;
                }
            }

            async Task CallApi()
            {
                try
                {
                    var apiConfig = ConfigManager.Config.ApiSettings;

                    using (HttpClient client = new HttpClient())
                    {
                        // headers from config
                        foreach (var header in apiConfig.RequestHeaders)
                        {
                            try
                            {
                                string headerValue = header.Value.Replace(
                                    GetConfigValue("ApiSettings.ApiKeyPlaceholder", "{API_KEY}"),
                                    apiConfig.ApiKey.Trim());
                                client.DefaultRequestHeaders.Remove(header.Key);
                                client.DefaultRequestHeaders.Add(header.Key, headerValue);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"{GetConfigValue("Messages.ErrorAddingHeader", "Error adding header")} {header.Key}: {ex.Message}");
                            }
                        }

                        client.Timeout = TimeSpan.FromSeconds(Math.Max(apiConfig.TimeoutSeconds,
                            GetConfigInt("ApiSettings.MinimumTimeoutSeconds", 30)));

                        // payload from config
                        var payloadDict = new Dictionary<string, object>();
                        foreach (var param in apiConfig.RequestParameters)
                        {
                            object value = param.Value;

                            if (value is string strValue)
                            {
                                value = strValue
                                    .Replace(GetConfigValue("ApiSettings.ModelNamePlaceholder", "{MODEL_NAME}"), apiConfig.ModelName ?? "")
                                    .Replace(GetConfigValue("ApiSettings.PromptPlaceholder", "{PROMPT}"), prompt ?? "");
                            }

                            if (param.Key == GetConfigValue("ApiSettings.MessagesKey", "messages") &&
                                value.ToString() == GetConfigValue("ApiSettings.MessagesPlaceholder", "{MESSAGES}"))
                            {
                                value = new[] { new { role = apiConfig.MessageRole, content = prompt } };
                            }

                            payloadDict[param.Key] = value;
                        }

                        string jsonPayload = JsonConvert.SerializeObject(payloadDict);
                        var httpMethod = new HttpMethod(GetConfigValue("ApiSettings.HttpMethod", "POST"));
                        var request = new HttpRequestMessage(httpMethod, apiConfig.ApiUrl)
                        {
                            Content = new StringContent(jsonPayload,
                                Encoding.GetEncoding(GetConfigValue("ApiSettings.ContentEncoding", "UTF-8")),
                                GetConfigValue("ApiSettings.ContentType", "application/json"))
                        };

                        using (var response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead))
                        {
                            if (!response.IsSuccessStatusCode)
                            {
                                string errorContent = await response.Content.ReadAsStringAsync();
                                var unauthorizedCode = GetConfigValue("ApiSettings.UnauthorizedStatusCode", "401");
                                if (response.StatusCode.ToString() == unauthorizedCode ||
                                    ((int)response.StatusCode).ToString() == unauthorizedCode)
                                    throw new Exception($"{ConfigManager.Config.Messages.ApiError} ({unauthorizedCode}): {errorContent}");
                                else
                                    throw new Exception($"{ConfigManager.Config.Messages.ApiError} ({response.StatusCode}): {errorContent}");
                            }

                            using (var stream = await response.Content.ReadAsStreamAsync())
                            using (var reader = new StreamReader(stream))
                            {
                                StringBuilder messageBuilder = new StringBuilder();
                                bool firstChunk = true;
                                string line;

                                while ((line = await reader.ReadLineAsync()) != null)
                                {
                                    if (string.IsNullOrWhiteSpace(line)) continue;

                                    string jsonLine = line;

                                    if (!string.IsNullOrEmpty(apiConfig.StreamPrefix) &&
                                        line.StartsWith(apiConfig.StreamPrefix))
                                    {
                                        jsonLine = line.Substring(apiConfig.StreamPrefix.Length).Trim();
                                    }

                                    if (!string.IsNullOrEmpty(apiConfig.StreamEndMarker) &&
                                        jsonLine == apiConfig.StreamEndMarker)
                                    {
                                        break;
                                    }

                                    try
                                    {
                                        var data = JsonConvert.DeserializeObject<dynamic>(jsonLine);

                                        // Prefer configured path; fallback to configured fallback paths
                                        string contentChunk = string.Empty;
                                        if (!string.IsNullOrWhiteSpace(apiConfig.ResponseContentPath))
                                        {
                                            contentChunk = ExtractContentFromResponse(data, apiConfig.ResponseContentPath);
                                        }

                                        if (string.IsNullOrEmpty(contentChunk) && config.ClassificationSettings?.ApiResponse?.FallbackContentPaths != null)
                                        {
                                            foreach (var fallbackPath in config.ClassificationSettings.ApiResponse.FallbackContentPaths)
                                            {
                                                contentChunk = ExtractContentFromResponse(data, fallbackPath);
                                                if (!string.IsNullOrEmpty(contentChunk)) break;
                                            }
                                        }

                                        if (!string.IsNullOrEmpty(contentChunk))
                                        {
                                            if (firstChunk)
                                            {
                                                SafeInvoke(loadingLabel, () => loadingLabel.Visible = false);
                                                SafeInvoke(textBox, () => textBox.Visible = true);
                                                dotTimer.Stop();
                                                firstChunk = false;
                                            }

                                            messageBuilder.Append(contentChunk);
                                            string currentMessage = messageBuilder.ToString();
                                            string displayText = isReplyMode
                                                ? ParseReplyCodeAndGetDisplayText(currentMessage)
                                                : currentMessage;

                                            SafeInvoke(textBox, () =>
                                            {
                                                textBox.Text = displayText;
                                                textBox.SelectionStart = textBox.Text.Length;
                                                textBox.ScrollToCaret();
                                            });

                                            await Task.Delay(config.UiSettings.StreamingSettings.CharacterDelay);
                                        }
                                    }
                                    catch (JsonException)
                                    {
                                        // ignore malformed event line
                                        continue;
                                    }
                                }

                                fullResponse = messageBuilder.ToString();
                                displayResponse = isReplyMode
                                    ? ParseReplyCodeAndGetDisplayText(fullResponse)
                                    : fullResponse;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    SafeInvoke(loadingLabel, () => loadingLabel.Visible = false);
                    SafeInvoke(textBox, () =>
                    {
                        textBox.Visible = true;
                        textBox.Clear();
                        textBox.AppendText($"{ConfigManager.Config.Messages.Error}: {ex.Message}\n\n{ConfigManager.Config.Messages.ClickRetry}");
                    });
                    SafeInvoke(retryButton, () => retryButton.Visible = true);
                }
                finally
                {
                    SafeInvoke(streamingForm, () =>
                    {
                        string completeSuffix = GetConfigValue("Messages.CompleteSuffix", " - Complete");
                        string dashSeparator = GetConfigValue("Messages.TitleSeparator", " - ");
                        streamingForm.Text = $"{windowTitle}{dashSeparator}{ConfigManager.Config.Messages.Complete}";
                        closeButton.Visible = true;

                        if (isReplyMode && replyCode != GetConfigInt("ReplySettings.DefaultReplyCode", -1))
                        {
                            if (replyCode == GetConfigInt("ReplySettings.NoReplyCode", 0))
                                generateAnywayButton.Visible = true;
                            else if (replyCode == GetConfigInt("ReplySettings.ReplyCode", 1))
                                sendReplyButton.Visible = true;
                        }
                    });
                }
            }

            // wire buttons
            sendReplyButton.Click += (s, ev) =>
            {
                try
                {
                    if (originalMail != null)
                    {
                        Outlook.MailItem reply = originalMail.Reply();
                        string replyTemplate = GetConfigValue("ReplySettings.ReplyBodyTemplate", "{RESPONSE}\n\n{ORIGINAL_BODY}");
                        reply.Body = replyTemplate
                            .Replace(GetConfigValue("ReplySettings.ResponsePlaceholder", "{RESPONSE}"), displayResponse)
                            .Replace(GetConfigValue("ReplySettings.OriginalBodyPlaceholder", "{ORIGINAL_BODY}"), reply.Body);
                        reply.Display();
                        streamingForm.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"{config.Messages.ErrorCreatingReply}: {ex.Message}");
                }
            };

            generateAnywayButton.Click += async (s, ev) =>
            {
                try
                {
                    string recentMessage = CleanMessage(ExtractMostRecentMessage(originalMail?.Body ?? ""));
                    string forceReplyPrompt = config.Prompts.ForceReply.Replace(
                        GetConfigValue("Prompts.MessagePlaceholder", "{MESSAGE}"), recentMessage);

                    SafeInvoke(loadingLabel, () => loadingLabel.Visible = true);
                    SafeInvoke(textBox, () => textBox.Visible = false);
                    SafeInvoke(generateAnywayButton, () => generateAnywayButton.Visible = false);
                    SafeInvoke(closeButton, () => closeButton.Visible = false);
                    SafeInvoke(streamingForm, () => streamingForm.Text = config.Messages.GeneratingReply);

                    replyCode = GetConfigInt("ReplySettings.ReplyCode", 1);
                    fullResponse = "";
                    displayResponse = "";
                    prompt = forceReplyPrompt;
                    dotTimer.Start();
                    await CallApi();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"{config.Messages.Error}: {ex.Message}");
                }
            };

            retryButton.Click += async (s, ev) =>
            {
                SafeInvoke(loadingLabel, () => loadingLabel.Visible = true);
                SafeInvoke(textBox, () => textBox.Visible = false);
                SafeInvoke(retryButton, () => retryButton.Visible = false);
                SafeInvoke(sendReplyButton, () => sendReplyButton.Visible = false);
                SafeInvoke(generateAnywayButton, () => generateAnywayButton.Visible = false);
                SafeInvoke(closeButton, () => closeButton.Visible = false);
                SafeInvoke(streamingForm, () => streamingForm.Text = windowTitle);

                replyCode = GetConfigInt("ReplySettings.DefaultReplyCode", -1);
                fullResponse = "";
                displayResponse = "";
                dotTimer.Start();
                await CallApi();
            };

            mainPanel.Controls.Add(loadingLabel);
            mainPanel.Controls.Add(textBox);
            streamingForm.Controls.Add(mainPanel);
            streamingForm.Controls.Add(bottomPanel);

            dotTimer.Start();

            streamingForm.Shown += async (s, e) =>
            {
                await CallApi();
            };

            // Use configured window display mode
            string displayMode = GetConfigValue("WindowSettings.MainWindow.DisplayMode", "Show");
            if (displayMode.ToLower() == "showdialog")
            {
                streamingForm.ShowDialog();
            }
            else
            {
                streamingForm.Show();
            }
        }

        private bool IsInitialized()
        {
            return _initialized && config != null;
        }

        // Updated button1_Click for Summarize - works from both Inspector and Explorer
        private async void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (!IsInitialized())
            {
                MessageBox.Show(GetConfigValue("Messages.AddinNotInitialized", "Add-in not properly initialized. Please check configuration and restart Outlook."),
                              GetConfigValue("Messages.InitializationErrorTitle", "Initialization Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Outlook.MailItem mail = GetCurrentMailItem();

                if (mail == null)
                {
                    MessageBox.Show(config.Messages.NoEmailSelected, config.Messages.NoEmailSelectedTitle,
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string body = mail.Body;
                if (string.IsNullOrWhiteSpace(body))
                {
                    MessageBox.Show(config.Messages.EmptyEmailBody, config.Messages.EmptyEmailTitle,
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string summarizePrompt = config.Prompts.Summarize.Replace(
                    GetConfigValue("Prompts.EmailContentPlaceholder", "{EMAIL_CONTENT}"), body);
                await StreamApiResponse(summarizePrompt, config.Messages.GeneratingSummary);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{config.Messages.ErrorGeneratingSummary}: {ex.Message}", config.Messages.Error,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Updated button2_Click for Translate - works from both Inspector and Explorer
        private async void button2_Click(object sender, RibbonControlEventArgs e)
        {
            if (!IsInitialized())
            {
                MessageBox.Show(GetConfigValue("Messages.AddinNotInitialized", "Add-in not properly initialized. Please check configuration and restart Outlook."),
                              GetConfigValue("Messages.InitializationErrorTitle", "Initialization Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Outlook.MailItem mail = GetCurrentMailItem();

                if (mail == null)
                {
                    MessageBox.Show(config.Messages.NoEmailSelected, config.Messages.NoEmailSelectedTitle,
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string body = mail.Body;
                if (string.IsNullOrWhiteSpace(body))
                {
                    MessageBox.Show(config.Messages.EmptyEmailBody, config.Messages.EmptyEmailTitle,
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string targetLanguage = string.IsNullOrWhiteSpace(config.LanguageSettings.DefaultTranslationLanguage)
                    ? config.LanguageSettings.FallbackLanguage
                    : config.LanguageSettings.DefaultTranslationLanguage;

                string translatePrompt = config.Prompts.Translate
                    .Replace(GetConfigValue("Prompts.EmailContentPlaceholder", "{EMAIL_CONTENT}"), body)
                    .Replace(GetConfigValue("Prompts.TargetLanguagePlaceholder", "{TARGET_LANGUAGE}"), targetLanguage);

                await StreamApiResponse(
                    translatePrompt,
                    config.Messages.Translating,
                    targetLanguage
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{config.Messages.ErrorTranslating}: {ex.Message}", config.Messages.Error,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Updated button3_Click for Smart Reply - works from both Inspector and Explorer
        private async void button3_Click(object sender, RibbonControlEventArgs e)
        {
            if (!IsInitialized())
            {
                MessageBox.Show(GetConfigValue("Messages.AddinNotInitialized", "Add-in not properly initialized. Please check configuration and restart Outlook."),
                              GetConfigValue("Messages.InitializationErrorTitle", "Initialization Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Outlook.MailItem mail = GetCurrentMailItem();

                if (mail == null)
                {
                    MessageBox.Show(config.Messages.NoEmailSelected, config.Messages.NoEmailSelectedTitle,
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string body = mail.Body;
                if (string.IsNullOrWhiteSpace(body))
                {
                    MessageBox.Show(config.Messages.EmptyEmailBody, config.Messages.EmptyEmailTitle,
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string recentMessage = CleanMessage(ExtractMostRecentMessage(body));
                string suggestReplyPrompt = config.Prompts.Reply.Replace(
                    GetConfigValue("Prompts.EmailContentPlaceholder", "{EMAIL_CONTENT}"), recentMessage);

                await StreamApiResponse(suggestReplyPrompt, config.Messages.AnalyzingEmail, null, true, mail);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{config.Messages.ErrorAnalyzing}: {ex.Message}", config.Messages.Error,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // NEW: button4_Click for Email Classification (merged from Ribbon2's button5)
        private async void button4_Click(object sender, RibbonControlEventArgs e)
        {
            if (!IsInitialized())
            {
                MessageBox.Show(
                    config.Messages.InitializationError ?? "Add-in not properly initialized",
                    config.Messages.ConfigurationError ?? "Configuration Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Check if classification is configured
            if (config.ClassificationSettings == null)
            {
                MessageBox.Show("Email classification feature is not configured. Please check your configuration file.",
                    "Feature Not Available", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                var application = Globals.ThisAddIn.Application;
                var nameSpace = application.GetNamespace(config.OutlookSettings?.MapiNamespace ?? "MAPI");
                var explorer = application.ActiveExplorer();

                Outlook.MAPIFolder currentFolder;

                // Use current folder if available, otherwise use configured default
                if (explorer != null && explorer.CurrentFolder != null)
                {
                    currentFolder = explorer.CurrentFolder;
                }
                else
                {
                    // Get folder type directly from config
                    string folderString = config.OutlookSettings?.DefaultFolders?.PrimaryFolder ?? "olFolderInbox";

                    Outlook.OlDefaultFolders folderType;
                    try
                    {
                        folderType = (Outlook.OlDefaultFolders)Enum.Parse(
                            typeof(Outlook.OlDefaultFolders),
                            folderString,
                            true);
                    }
                    catch
                    {
                        // Fallback to configured fallback folder
                        string fallbackFolder = config.OutlookSettings?.DefaultFolders?.FallbackFolder ?? "olFolderInbox";
                        try
                        {
                            folderType = (Outlook.OlDefaultFolders)Enum.Parse(
                                typeof(Outlook.OlDefaultFolders),
                                fallbackFolder,
                                true);
                        }
                        catch
                        {
                            // last fallback
                            folderType = Outlook.OlDefaultFolders.olFolderInbox;
                        }
                    }

                    currentFolder = nameSpace.GetDefaultFolder(folderType);
                }

                // Use configured unread filter
                string unreadFilter = config.OutlookSettings?.ItemRestrictions?.UnreadFilter ?? "[Unread] = true";

                var unreadItems = currentFolder.Items.Restrict(unreadFilter);
                var unreadEmails = new List<Outlook.MailItem>();

                foreach (object it in unreadItems)
                {
                    var mail = it as Outlook.MailItem;
                    if (mail != null) unreadEmails.Add(mail);
                }

                if (unreadEmails.Count == 0)
                {
                    string noUnread = FormatTemplate(
                        config.ClassificationSettings.Messages?.NoUnreadEmails ?? "No unread emails found in {FOLDER}",
                        new Dictionary<string, string>
                        {
                            {"{FOLDER}", currentFolder.Name ?? "" }
                        });

                    MessageBox.Show(noUnread,
                        config.Messages.NoEmailSelectedTitle ?? "No Email Selected",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Check if AI and rule-based classification are configured
                string aiStatus = (config.ClassificationSettings.AiClassification?.EnableAiClassification == true) ? "Enabled" : "Disabled";
                string ruleStatus = (config.ClassificationSettings.PreProcessingRules?.EnableRuleBasedClassification == true) ? "Enabled" : "Disabled";

                string confirmText = FormatTemplate(
                    config.ClassificationSettings.Messages?.ConfirmClassification ?? "Process {COUNT} unread emails in {FOLDER}?\n\nAI Classification: {AI_STATUS}\nRule-based Classification: {RULE_STATUS}",
                    new Dictionary<string, string>
                    {
                        {"{COUNT}", unreadEmails.Count.ToString()},
                        {"{FOLDER}", currentFolder.Name ?? ""},
                        {"{AI_STATUS}", aiStatus},
                        {"{RULE_STATUS}", ruleStatus}
                    });

                var confirm = MessageBox.Show(confirmText,
                    config.Messages.AnalyzingEmail ?? "Analyzing Email",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (confirm != DialogResult.Yes) return;

                // Prepare batch run
                ResetRateLimitingState();
                CreateCategoriesIfNeeded(application.Session);

                ShowProgressForm(unreadEmails.Count);
                int processed = 0, successful = 0, failed = 0;

                foreach (var mail in unreadEmails)
                {
                    try
                    {
                        processed++;

                        string subjectPreview = mail.Subject ?? "";
                        int previewLength = config.ClassificationSettings.Debug?.ContentPreviewLength ?? 100;
                        string truncationIndicator = config.ClassificationSettings.Debug?.TruncationIndicator ?? "...";

                        if (subjectPreview.Length > previewLength)
                            subjectPreview = subjectPreview.Substring(0, previewLength) + truncationIndicator;

                        string statusTemplate = config.ClassificationSettings.Progress?.StatusTemplate ?? "{CURRENT}/{TOTAL}: {SUBJECT}";

                        string statusText = FormatTemplate(
                            statusTemplate,
                            new Dictionary<string, string>
                            {
                                {"{CURRENT}", processed.ToString()},
                                {"{TOTAL}", unreadEmails.Count.ToString()},
                                {"{SUBJECT}", subjectPreview}
                            });

                        if (_isPaused && DateTime.Now < _pauseUntil)
                        {
                            TimeSpan waitRemain = _pauseUntil - DateTime.Now;
                            string timeFormat = config.ClassificationSettings.Progress?.TimeFormatString ?? "mm\\:ss";
                            string cooldownTemplate = config.ClassificationSettings.Progress?.CooldownTemplate ?? "  [Cooldown {TIME}]";

                            string cooldownText = FormatTemplate(cooldownTemplate,
                                new Dictionary<string, string>
                                {
                                    {"{TIME}", waitRemain.ToString(timeFormat)}
                                });
                            statusText += cooldownText;
                        }

                        UpdateProgress(statusText, processed, unreadEmails.Count);
                        await ApplyRateLimit().ConfigureAwait(false);

                        var info = GetEmailInfo(mail);
                        var classificationResult = await ClassifyEmailEnhanced(info).ConfigureAwait(false);

                        if (classificationResult.Success)
                        {
                            ApplyClassificationWithVisualIndicators(mail, classificationResult.ClassificationKey);
                            successful++;
                            _consecutiveFailures = 0;
                        }
                        else
                        {
                            ApplyClassificationWithVisualIndicators(mail, classificationResult.ClassificationKey);
                            failed++;
                        }
                    }
                    catch (Exception ex)
                    {
                        failed++;
                        System.Diagnostics.Debug.WriteLine($"Error processing email '{mail?.Subject}': {ex.Message}");
                    }
                }

                CloseProgressForm();

                // Reset view using config setting
                if (config.OutlookSettings?.ViewSettings?.ResetViewAfterClassification == true)
                {
                    try { explorer?.CurrentView?.Reset(); } catch { /* ignore */ }
                }

                string summary = FormatTemplate(
                    config.ClassificationSettings.Messages?.CompletionSummary ?? "Classification complete.\n\nProcessed: {PROCESSED}\nSuccessful: {SUCCESS}\nFailed: {FAILED}",
                    new Dictionary<string, string>
                    {
                        {"{PROCESSED}", processed.ToString()},
                        {"{SUCCESS}", successful.ToString()},
                        {"{FAILED}", failed.ToString()}
                    });

                MessageBox.Show(summary,
                    config.Messages.Complete ?? "Complete",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                CloseProgressForm();
                MessageBox.Show($"{config.Messages.Error ?? "Error"}: {ex.Message}",
                    config.Messages.Error ?? "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Helper method to get the current mail item from either source
        private Outlook.MailItem GetCurrentMailItem()
        {
            // First, try to get email from ActiveInspector (opened email window)
            var inspector = Globals.ThisAddIn.Application.ActiveInspector();
            if (inspector != null && inspector.CurrentItem is Outlook.MailItem inspectorMail)
            {
                return inspectorMail;
            }

            // If no inspector, try to get selected email from ActiveExplorer (main window)
            var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            if (explorer != null && explorer.Selection.Count > 0)
            {
                return explorer.Selection[1] as Outlook.MailItem; // Outlook collections are 1-indexed
            }

            return null;
        }
        private string FormatTemplate(string template, Dictionary<string, string> values)
        {
            if (string.IsNullOrEmpty(template)) return string.Empty;
            if (values == null || values.Count == 0) return template;

            string result = template;
            foreach (var kv in values)
            {
                string key = kv.Key;
                if (!key.StartsWith("{")) key = "{" + key + "}";
                result = result.Replace(key, kv.Value ?? "");
            }

            return result;
        }

        private class EmailInfo
        {
            public string Subject;
            public string Sender;
            public string Body;
        }

        private EmailInfo GetEmailInfo(Outlook.MailItem mail)
        {
            var emailProc = config.ClassificationSettings.EmailProcessing;

            string subject = mail.Subject ?? "";
            string sender = mail.SenderEmailAddress ?? mail.SenderName ?? "";
            string body = mail.Body ?? "";

            // Clean the subject of previous AI classifications
            subject = CleanSubjectFromAITags(subject);

            // Trim to the most recent message using separators
            var separators = new List<string>();
            if (emailProc.ContentSeparators != null) separators.AddRange(emailProc.ContentSeparators);
            if (config.EmailSettings?.MessageSeparators != null)
                separators.AddRange(config.EmailSettings.MessageSeparators);

            // Add common separators if none configured
            if (separators.Count == 0)
            {
                separators.AddRange(new[] { "-----Original Message-----", "From:", "On ", "________________________________" });
            }

            string originalBody = body;
            int firstSeparator = -1;
            foreach (var sep in separators)
            {
                if (string.IsNullOrWhiteSpace(sep)) continue;
                int pos = body.IndexOf(sep, StringComparison.OrdinalIgnoreCase);
                if (pos > 0 && (firstSeparator == -1 || pos < firstSeparator))
                    firstSeparator = pos;
            }
            if (firstSeparator > 0)
            {
                body = body.Substring(0, firstSeparator).Trim();
            }

            // Remove signatures
            if (config.EmailSettings?.SignaturePrefixes != null &&
                config.EmailSettings.SignaturePrefixes.Count > 0)
            {
                var lines = body.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                var contentLines = new List<string>();

                foreach (var line in lines)
                {
                    bool isSignature = config.EmailSettings.SignaturePrefixes.Any(prefix =>
                        line.Trim().StartsWith(prefix, StringComparison.OrdinalIgnoreCase));

                    if (isSignature) break;
                    contentLines.Add(line);
                }

                body = string.Join("\n", contentLines).Trim();
            }

            // If body is empty after processing, try to use some minimal content
            if (string.IsNullOrWhiteSpace(body) && !string.IsNullOrWhiteSpace(originalBody))
            {
                body = originalBody.Trim();
                if (body.Length > 200) body = body.Substring(0, 200) + "...";
            }

            // If STILL empty, create minimal content from subject
            if (string.IsNullOrWhiteSpace(body))
            {
                body = FormatTemplate(config.ClassificationSettings.Content?.FallbackContentTemplate ?? "Subject: {SUBJECT}",
                    new Dictionary<string, string> { { "{SUBJECT}", subject } });
            }

            // Limit body length
            int maxLen = emailProc.MaxBodyLength > 0 ? emailProc.MaxBodyLength : 800;
            if (body.Length > maxLen)
            {
                body = body.Substring(0, maxLen) + "...";
            }

            return new EmailInfo
            {
                Subject = subject.Trim(),
                Sender = sender.Trim(),
                Body = body.Trim()
            };
        }

        private string CleanSubjectFromAITags(string subject)
        {
            if (string.IsNullOrWhiteSpace(subject)) return subject;

            // Remove various AI classification prefixes
            var aiPrefixes = config.ClassificationSettings.Content?.AiTagPatterns ?? new List<string>();

            string cleaned = subject;
            foreach (var prefix in aiPrefixes)
            {
                cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, prefix, "",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            }

            return cleaned.Trim();
        }

        private async Task<(bool Success, string ClassificationKey)> ClassifyEmailEnhanced(EmailInfo info)
        {
            // STEP 1: Try rule-based classification first (if enabled and rules exist)
            if (config.ClassificationSettings.PreProcessingRules?.EnableRuleBasedClassification == true &&
                config.ClassificationSettings.PreProcessingRules.Rules?.Count > 0)
            {
                string ruleBasedResult = ApplyRuleBasedClassification(info);
                if (!string.IsNullOrEmpty(ruleBasedResult))
                {
                    System.Diagnostics.Debug.WriteLine($"Email classified by rule: {ruleBasedResult}");
                    return (true, ruleBasedResult);
                }
            }

            // STEP 2: Use AI classification if no rules matched or AI is enabled
            if (config.ClassificationSettings.AiClassification?.EnableAiClassification == true)
            {
                bool useAiForUnmatched = config.ClassificationSettings.AiClassification.UseAiForUnmatched;

                if (useAiForUnmatched ||
                    config.ClassificationSettings.PreProcessingRules?.EnableRuleBasedClassification != true ||
                    config.ClassificationSettings.PreProcessingRules?.Rules?.Count == 0)
                {
                    return await ClassifyEmailWithRetry(info).ConfigureAwait(false);
                }
            }

            // STEP 3: Fallback to default
            return (false, GetDefaultClassification());
        }

        private string ApplyRuleBasedClassification(EmailInfo info)
        {
            var rules = config.ClassificationSettings.PreProcessingRules?.Rules;
            if (rules == null || rules.Count == 0) return null;

            // Sort rules by priority (highest first)
            var sortedRules = rules.OrderByDescending(r => r.Priority).ToList();

            foreach (var rule in sortedRules)
            {
                if (EvaluateRule(rule, info))
                {
                    System.Diagnostics.Debug.WriteLine($"Rule '{rule.Name}' matched - Classification: {rule.Classification}");
                    return rule.Classification;
                }
            }

            return null; // No rules matched
        }

        private bool EvaluateRule(ClassificationRule rule, EmailInfo info)
        {
            if (rule?.Conditions == null) return false;

            var conditions = rule.Conditions;
            var matches = new List<bool>();

            // Check subject keywords
            if (conditions.SubjectKeywords != null && conditions.SubjectKeywords.Count > 0)
            {
                bool subjectMatch = conditions.SubjectKeywords.Any(keyword =>
                    (info.Subject ?? "").IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);
                matches.Add(subjectMatch);
            }

            // Check sender domains
            if (conditions.SenderDomains != null && conditions.SenderDomains.Count > 0)
            {
                bool domainMatch = conditions.SenderDomains.Any(domain =>
                {
                    string sender = info.Sender ?? "";
                    return sender.EndsWith("@" + domain, StringComparison.OrdinalIgnoreCase) ||
                           sender.IndexOf(domain, StringComparison.OrdinalIgnoreCase) >= 0;
                });
                matches.Add(domainMatch);
            }

            // Check sender addresses
            if (conditions.SenderAddresses != null && conditions.SenderAddresses.Count > 0)
            {
                bool addressMatch = conditions.SenderAddresses.Any(address =>
                    string.Equals(info.Sender ?? "", address, StringComparison.OrdinalIgnoreCase));
                matches.Add(addressMatch);
            }

            // Check body keywords
            if (conditions.BodyKeywords != null && conditions.BodyKeywords.Count > 0)
            {
                bool bodyMatch = conditions.BodyKeywords.Any(keyword =>
                    (info.Body ?? "").IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);
                matches.Add(bodyMatch);
            }

            // Apply match type logic
            if (matches.Count == 0) return false;

            bool result = conditions.MatchType?.ToUpperInvariant() == "ALL"
                ? matches.All(m => m)  // ALL conditions must match
                : matches.Any(m => m); // ANY condition must match

            System.Diagnostics.Debug.WriteLine($"Rule '{rule.Name}': {matches.Count} conditions, result: {result}");
            return result;
        }

        // Add these rate limiting fields to your class variables
        private DateTime _lastRequestTime = DateTime.MinValue;
        private readonly Queue<DateTime> _requestHistory = new Queue<DateTime>();
        private int _consecutiveFailures = 0;
        private bool _isPaused = false;
        private DateTime _pauseUntil = DateTime.MinValue;
        private ProgressForm _progressForm;

        private void ResetRateLimitingState()
        {
            _consecutiveFailures = 0;
            _isPaused = false;
            _pauseUntil = DateTime.MinValue;
            _requestHistory.Clear();
            _lastRequestTime = DateTime.MinValue;
        }

        private async Task ApplyRateLimit()
        {
            var rl = config.ClassificationSettings.RateLimiting;

            // Active cooldown
            if (_isPaused && DateTime.Now < _pauseUntil)
            {
                var wait = _pauseUntil - DateTime.Now;
                await WaitWithProgress(wait, "Cooldown").ConfigureAwait(false);
                _isPaused = false;
            }

            // Clean history older than 1 minute
            var cutoff = DateTime.Now.AddMinutes(-1);
            while (_requestHistory.Count > 0 && _requestHistory.Peek() < cutoff)
            {
                _requestHistory.Dequeue();
            }

            // RPM guard
            if (rl.RequestsPerMinute > 0 && _requestHistory.Count >= rl.RequestsPerMinute)
            {
                var oldest = _requestHistory.Peek();
                var waitUntil = oldest.AddMinutes(1);
                var wait = waitUntil - DateTime.Now;
                if (wait.TotalMilliseconds > 0)
                    await WaitWithProgress(wait, "Rate limit").ConfigureAwait(false);
            }

            // Min delay between requests
            int minDelayMs = rl.MinDelayBetweenRequestsMs > 0 ? rl.MinDelayBetweenRequestsMs : 1500;
            var sinceLast = DateTime.Now - _lastRequestTime;
            if (sinceLast.TotalMilliseconds < minDelayMs)
            {
                var remaining = TimeSpan.FromMilliseconds(minDelayMs) - sinceLast;
                if (remaining.TotalMilliseconds > 0)
                    await Task.Delay(remaining).ConfigureAwait(false);
            }

            _lastRequestTime = DateTime.Now;
        }

        private async Task WaitWithProgress(TimeSpan waitTime, string reason)
        {
            DateTime end = DateTime.Now.Add(waitTime);

            while (DateTime.Now < end)
            {
                var remaining = end - DateTime.Now;
                if (remaining.TotalSeconds <= 0) break;

                string text = FormatTemplate(config.ClassificationSettings.Messages.WaitStatus,
                    new Dictionary<string, string>
                    {
                {"{REASON}", reason},
                {"{TIME}", remaining.ToString(config.ClassificationSettings.Progress.TimeFormatString)}
                    });

                UpdateProgress(text, -1, -1); // -1 means don't update progress bar

                await Task.Delay(1000).ConfigureAwait(false);
            }
        }

        private async Task<(bool Success, string ClassificationKey)> ClassifyEmailWithRetry(EmailInfo info)
        {
            var rl = config.ClassificationSettings.RateLimiting;
            var retry = config.ClassificationSettings.Retry;
            int maxAttempts = retry.MaxAttempts;
            Exception lastException = null;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    _requestHistory.Enqueue(DateTime.Now);
                    string key = await ClassifyEmailOnce(info).ConfigureAwait(false);

                    if (!string.IsNullOrEmpty(key) &&
                        config.ClassificationSettings.Classifications.ContainsKey(key))
                    {
                        return (true, key);
                    }

                    // Invalid classification → force retry
                    System.Diagnostics.Debug.WriteLine($"Invalid classification returned: '{key}'");
                    throw new Exception($"Invalid classification result: '{key}'");
                }
                catch (HttpRequestException ex) when (IsRateLimitException(ex))
                {
                    lastException = ex;
                    _consecutiveFailures++;
                    await HandleRateLimit(ex, attempt).ConfigureAwait(false);
                }
                catch (TaskCanceledException ex)
                {
                    lastException = ex;
                    _consecutiveFailures++;
                    System.Diagnostics.Debug.WriteLine($"Request timeout on attempt {attempt}: {ex.Message}");

                    if (attempt < maxAttempts)
                    {
                        int baseTimeout = retry.BackoffMultipliers?.TimeoutRetryBase > 0
                            ? retry.BackoffMultipliers.TimeoutRetryBase
                            : 5; // fallback to old behavior

                        int expBase = retry.BackoffMultipliers?.ExponentialBase > 0
                            ? retry.BackoffMultipliers.ExponentialBase
                            : 2; // default doubling

                        var delay = TimeSpan.FromSeconds(baseTimeout * Math.Pow(expBase, attempt));
                        await WaitWithProgress(delay, $"Timeout retry {attempt + 1}/{maxAttempts}").ConfigureAwait(false);
                    }
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    _consecutiveFailures++;
                    System.Diagnostics.Debug.WriteLine($"Classification error attempt {attempt}: {ex.Message}");

                    if (attempt < maxAttempts)
                    {
                        int baseError = retry.BackoffMultipliers?.ErrorRetryBase > 0
                            ? retry.BackoffMultipliers.ErrorRetryBase
                            : (rl.BaseCooldownSeconds > 0 ? rl.BaseCooldownSeconds : 10);

                        int expBase = retry.BackoffMultipliers?.ExponentialBase > 0
                            ? retry.BackoffMultipliers.ExponentialBase
                            : 2;

                        var delay = TimeSpan.FromSeconds(baseError * Math.Pow(expBase, attempt));
                        await WaitWithProgress(delay, $"Error retry {attempt + 1}/{maxAttempts}").ConfigureAwait(false);
                    }
                }

                // Extended cooldown after too many consecutive failures
                // Extended cooldown after consecutive failures
                if (_consecutiveFailures >= (rl.MaxConsecutiveFailures > 0 ? rl.MaxConsecutiveFailures : 3))
                {
                    _isPaused = true;

                    int factor = config.ClassificationSettings.Retry?.BackoffMultipliers?.ExtendedCooldownFactor > 0
                        ? config.ClassificationSettings.Retry.BackoffMultipliers.ExtendedCooldownFactor
                        : 6; // default

                    int cooldownSec = rl.BaseCooldownSeconds > 0 ? rl.BaseCooldownSeconds * factor : 60;
                    int maxMin = rl.MaxCooldownMinutes > 0 ? rl.MaxCooldownMinutes : 5;
                    var extended = TimeSpan.FromSeconds(Math.Min(cooldownSec, maxMin * 60));
                    _pauseUntil = DateTime.Now.Add(extended);

                    await WaitWithProgress(extended, $"Extended cooldown ({_consecutiveFailures} failures)").ConfigureAwait(false);
                    _consecutiveFailures = 0;
                }

            }

            // Final failure
            System.Diagnostics.Debug.WriteLine($"All classification attempts failed. Last error: {lastException?.Message}");
            return (false, GetDefaultClassification());
        }


        private bool IsRateLimitException(HttpRequestException ex)
        {
            if (ex?.Message == null) return false;

            return ex.Message.Contains("429") ||
                   ex.Message.IndexOf("rate limit", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   ex.Message.IndexOf("too many requests", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private string GetDefaultClassification()
        {
            // UPDATED: Use flexible fallback order from config
            var fallbacks = config.ClassificationSettings.Retry.DefaultFallbackOrder;

            foreach (var fallback in fallbacks)
            {
                if (config.ClassificationSettings.Classifications.ContainsKey(fallback))
                    return fallback;
            }

            // Last resort - first configured classification (any classification is valid now)
            return config.ClassificationSettings.Classifications.Keys.FirstOrDefault() ?? "DEFAULT";
        }

        private async Task HandleRateLimit(Exception ex, int attempt)
        {
            var rl = config.ClassificationSettings.RateLimiting;

            int baseSec = rl.BaseCooldownSeconds > 0 ? rl.BaseCooldownSeconds : 10;
            int maxMin = rl.MaxCooldownMinutes > 0 ? rl.MaxCooldownMinutes : 5;

            // Exponential backoff based on base cooldown
            var backoff = TimeSpan.FromSeconds(Math.Pow(2, attempt) * baseSec);
            var baseWait = TimeSpan.FromSeconds(baseSec);
            var waitTime = baseWait + backoff;

            // Cap
            var maxDelay = TimeSpan.FromMinutes(maxMin);
            if (waitTime > maxDelay) waitTime = maxDelay;

            _isPaused = true;
            _pauseUntil = DateTime.Now.Add(waitTime);

            await WaitWithProgress(waitTime, $"Rate limit (attempt {attempt})").ConfigureAwait(false);
        }

        private async Task<string> ClassifyEmailOnce(EmailInfo info)
        {
            string prompt = BuildClassificationPrompt(info);
            var api = config.ApiSettings;
            var rl = config.ClassificationSettings.RateLimiting;

            using (var client = new HttpClient())
            {
                try
                {
                    // Headers from config
                    if (api.RequestHeaders != null)
                    {
                        foreach (var header in api.RequestHeaders)
                        {
                            try
                            {
                                string headerValue = (header.Value ?? "").Replace("{API_KEY}", api.ApiKey.Trim());
                                client.DefaultRequestHeaders.Remove(header.Key);
                                client.DefaultRequestHeaders.Add(header.Key, headerValue);
                            }
                            catch (Exception headerEx)
                            {
                                System.Diagnostics.Debug.WriteLine($"Failed to add header {header.Key}: {headerEx.Message}");
                            }
                        }
                    }

                    int timeoutSeconds = rl.RequestTimeoutSeconds > 0 ? rl.RequestTimeoutSeconds :
                                         (api.TimeoutSeconds > 0 ? api.TimeoutSeconds : 45);
                    client.Timeout = TimeSpan.FromSeconds(timeoutSeconds);

                    // Build payload from RequestParameters (same pattern as Ribbon1)
                    var payloadDict = new Dictionary<string, object>();
                    if (api.RequestParameters != null)
                    {
                        foreach (var kv in api.RequestParameters)
                        {
                            object value = kv.Value;

                            if (value is string s)
                            {
                                value = (s ?? "")
                                    .Replace("{MODEL_NAME}", api.ModelName ?? "")
                                    .Replace("{PROMPT}", prompt ?? "");
                            }

                            if (kv.Key == "messages" && value is string placeholder && placeholder == "{MESSAGES}")
                            {
                                // Use message role from config, content = prompt
                                value = new[] { new { role = api.MessageRole ?? "user", content = prompt } };
                            }

                            payloadDict[kv.Key] = value;
                        }
                    }

                    // For classification, we don't need streaming
                    if (payloadDict.ContainsKey("stream"))
                    {
                        payloadDict["stream"] = false;
                    }

                    var jsonPayload = JsonConvert.SerializeObject(payloadDict);
                    System.Diagnostics.Debug.WriteLine($"API Request URL: {api.ApiUrl}");
                    System.Diagnostics.Debug.WriteLine($"API Request Payload: {jsonPayload}");

                    var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                    using (var response = await client.PostAsync(api.ApiUrl, content).ConfigureAwait(false))
                    {
                        string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        System.Diagnostics.Debug.WriteLine($"API Response Status: {response.StatusCode}");
                        System.Diagnostics.Debug.WriteLine($"API Response Body: {body}");

                        // NOTE: HttpStatusCode.TooManyRequests may not exist in older frameworks, use numeric 429 check
                        if ((int)response.StatusCode == 429)
                        {
                            throw new HttpRequestException("429 Too Many Requests - Rate limited");
                        }

                        if (!response.IsSuccessStatusCode)
                        {
                            throw new HttpRequestException($"API Error {(int)response.StatusCode} {response.StatusCode}: {body}");
                        }

                        // Non-streaming parse using ResponseContentPath (same helper as Ribbon1)
                        dynamic data = null;
                        try
                        {
                            data = JsonConvert.DeserializeObject<dynamic>(body);
                        }
                        catch (JsonException jsonEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"JSON parse error: {jsonEx.Message}");
                            throw new Exception($"Invalid JSON response: {jsonEx.Message}");
                        }

                        string text = ExtractContentFromResponse(data, api.ResponseContentPath);
                        if (string.IsNullOrWhiteSpace(text))
                        {
                            // Fallback: try common response paths
                            var fallbackPaths = config.ClassificationSettings.ApiResponse.FallbackContentPaths;
                            foreach (var path in fallbackPaths)
                            {
                                text = ExtractContentFromResponse(data, path);
                                if (!string.IsNullOrWhiteSpace(text)) break;
                            }
                        }

                        if (string.IsNullOrWhiteSpace(text))
                        {
                            throw new Exception($"No content found in API response using path: {api.ResponseContentPath}");
                        }

                        System.Diagnostics.Debug.WriteLine($"Extracted classification text: '{text}'");

                        // Normalize classification text to one of the configured keys
                        return NormalizeClassificationKey(text);
                    }
                }
                catch (TaskCanceledException ex)
                {
                    throw new TaskCanceledException($"API request timeout after {client.Timeout.TotalSeconds}s: {ex.Message}", ex);
                }
                catch (HttpRequestException ex)
                {
                    throw new HttpRequestException($"HTTP request failed: {ex.Message}", ex);
                }
            }
        }

        private string BuildClassificationPrompt(EmailInfo info)
        {
            if (config.ClassificationSettings.AiClassification?.DynamicPrompts?.IncludeAvailableCategories == true)
            {
                return BuildDynamicClassificationPrompt(info);
            }

            // Get the base prompt template
            string prompt = config.ClassificationSettings.Prompt;

            // Create a comprehensive email content string
            string emailContent = $"Subject: {info.Subject ?? "[No Subject]"}\nSender: {info.Sender ?? "[No Sender]"}\nContent: {info.Body ?? "[No Body]"}";

            // Replace placeholders with actual email content
            prompt = prompt.Replace("{SUBJECT}", info.Subject ?? "[No Subject]");
            prompt = prompt.Replace("{SENDER}", info.Sender ?? "[No Sender]");
            prompt = prompt.Replace("{BODY}", info.Body ?? "[No Body]");
            prompt = prompt.Replace("{EMAIL_CONTENT}", emailContent);
            prompt = prompt.Replace("{CONTENT}", emailContent);
            prompt = prompt.Replace("{EMAIL}", emailContent);

            return prompt;
        }

        private string BuildDynamicClassificationPrompt(EmailInfo info)
        {
            var aiSettings = config.ClassificationSettings.AiClassification;
            if (aiSettings?.DynamicPrompts?.IncludeAvailableCategories != true)
            {
                return BuildClassificationPrompt(info); // Use original method
            }

            var dynamicPrompts = aiSettings.DynamicPrompts;
            string template = dynamicPrompts.PromptTemplate ?? config.ClassificationSettings.Prompt;

            // Get available categories
            var availableCategories = string.Join(", ", config.ClassificationSettings.Classifications.Keys);

            // Build category descriptions
            var descriptions = new List<string>();
            foreach (var kv in config.ClassificationSettings.Classifications)
            {
                string description = "";
                if (dynamicPrompts.CategoryDescriptions?.ContainsKey(kv.Key) == true)
                {
                    description = dynamicPrompts.CategoryDescriptions[kv.Key];
                }
                else
                {
                    description = $"{kv.Key} emails";
                }
                descriptions.Add($"- {kv.Key}: {description}");
            }
            string categoryDescriptions = string.Join("\n", descriptions);

            // Create email content
            string emailContent = $"Subject: {info.Subject ?? "[No Subject]"}\nSender: {info.Sender ?? "[No Sender]"}\nContent: {info.Body ?? "[No Body]"}";

            // Replace placeholders
            string prompt = template
                .Replace("{AVAILABLE_CATEGORIES}", availableCategories)
                .Replace("{CATEGORY_DESCRIPTIONS}", categoryDescriptions)
                .Replace("{EMAIL_CONTENT}", emailContent)
                .Replace("{SUBJECT}", info.Subject ?? "[No Subject]")
                .Replace("{SENDER}", info.Sender ?? "[No Sender]")
                .Replace("{BODY}", info.Body ?? "[No Body]");

            System.Diagnostics.Debug.WriteLine($"Dynamic Prompt Generated: {prompt}");
            return prompt;
        }

        private string ExtractContentFromResponse(dynamic data, string path)
        {
            try
            {
                if (data == null) return string.Empty;
                if (string.IsNullOrWhiteSpace(path)) return data.ToString();

                var pathParts = path.Split(new[] { '[', ']', '.' }, StringSplitOptions.RemoveEmptyEntries);
                object current = data;

                foreach (var part in pathParts)
                {
                    if (current == null) break;

                    var jObject = current as Newtonsoft.Json.Linq.JObject;
                    if (jObject != null)
                    {
                        current = jObject[part];
                        continue;
                    }

                    var jArray = current as Newtonsoft.Json.Linq.JArray;
                    int index;
                    if (jArray != null && int.TryParse(part, out index))
                    {
                        current = index < jArray.Count ? jArray[index] : null;
                        continue;
                    }

                    current = null;
                }

                return current != null ? current.ToString() : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string NormalizeClassificationKey(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return GetDefaultClassification();

            string text = raw.Trim().ToUpperInvariant();
            System.Diagnostics.Debug.WriteLine($"Normalizing classification: '{text}'");

            // Handle the specific response you're getting
            bool hasNoContentIndicator = config.ClassificationSettings.ApiResponse.NoContentIndicators
                .Any(indicator => text.Contains(indicator));
            if (hasNoContentIndicator)
            {
                System.Diagnostics.Debug.WriteLine("Detected 'no content' response - prompt issue!");
                return GetDefaultClassification();
            }

            // Exact match first
            foreach (var key in config.ClassificationSettings.Classifications.Keys)
            {
                if (string.Equals(text, key, StringComparison.OrdinalIgnoreCase))
                {
                    System.Diagnostics.Debug.WriteLine($"Exact match found: '{key}'");
                    return key;
                }
            }

            // Contains match (e.g., "The email is PANCAKES.")
            foreach (var key in config.ClassificationSettings.Classifications.Keys)
            {
                if (text.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    System.Diagnostics.Debug.WriteLine($"Contains match found: '{key}' in '{text}'");
                    return key;
                }
            }

            System.Diagnostics.Debug.WriteLine($"No match found, using default classification");
            return GetDefaultClassification();
        }

        private void ApplyClassificationWithVisualIndicators(Outlook.MailItem mail, string classificationKey)
        {
            try
            {
                ClassificationConfig cfg;
                if (!config.ClassificationSettings.Classifications.TryGetValue(classificationKey, out cfg) || cfg == null)
                    return;

                // Category name
                string categoryName = cfg.CategoryPrefix ?? classificationKey;

                // Outlook importance
                Outlook.OlImportance importance = ParseImportance(cfg.Importance);

                // Flag icon
                Outlook.OlFlagIcon flagIcon;
                bool hasFlagIcon = TryParseFlagIcon(cfg.FlagIcon, out flagIcon);

                // Flag request text and subject prefix
                string flagRequest = cfg.FlagRequest ?? "";
                string subjectPrefix = cfg.SubjectPrefix ?? "";

                // Apply importance + flag
                mail.Importance = importance;

                if (hasFlagIcon)
                {
                    mail.FlagStatus = Outlook.OlFlagStatus.olFlagMarked;
                    mail.FlagIcon = flagIcon;
                }
                else
                {
                    mail.FlagStatus = Outlook.OlFlagStatus.olNoFlag;
                }

                if (!string.IsNullOrEmpty(flagRequest))
                {
                    mail.FlagRequest = flagRequest;
                }

                // FIXED: Categories - properly clean and set categories
                string currentCategories = mail.Categories ?? "";
                var categoryList = new List<string>();

                // Parse existing categories
                if (!string.IsNullOrEmpty(currentCategories))
                {
                    categoryList = currentCategories
                        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(c => c.Trim())
                        .Where(c => !string.IsNullOrEmpty(c))
                        .ToList();
                }

                // UPDATED: Remove classification-related categories more flexibly
                var classificationPrefixes = new List<string>(config.ClassificationSettings.Content.ClassificationKeywords);

                // Add ALL configured category prefixes and keys (not hardcoded anymore)
                foreach (var kv in config.ClassificationSettings.Classifications)
                {
                    if (kv.Value != null && !string.IsNullOrEmpty(kv.Value.CategoryPrefix))
                        classificationPrefixes.Add(kv.Value.CategoryPrefix.ToUpperInvariant());
                    classificationPrefixes.Add(kv.Key.ToUpperInvariant());
                }

                // Remove any category that contains classification keywords
                categoryList = categoryList
                    .Where(c => !classificationPrefixes.Any(prefix =>
                        c.ToUpperInvariant().Contains(prefix)))
                    .ToList();

                // Add ONLY the new category (avoid duplicates)
                if (!string.IsNullOrEmpty(categoryName) &&
                    !categoryList.Any(c => c.Equals(categoryName, StringComparison.OrdinalIgnoreCase)))
                {
                    categoryList.Add(categoryName);
                }

                // Set the cleaned categories
                mail.Categories = string.Join(",", categoryList.Distinct());

                // FIXED: Subject prefix handling - avoid duplicates and clean existing AI tags first
                if (!string.IsNullOrWhiteSpace(subjectPrefix))
                {
                    try
                    {
                        string currentSubject = mail.Subject ?? "";

                        // FIRST: Clean existing AI tags from the subject
                        currentSubject = CleanSubjectFromAITags(currentSubject);

                        // THEN: Check if our new prefix already exists (case insensitive)
                        if (!currentSubject.StartsWith(subjectPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            mail.Subject = $"{subjectPrefix} {currentSubject}".Trim();
                        }
                        else
                        {
                            // Prefix already exists, just update the existing subject
                            mail.Subject = currentSubject;
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException comEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"Cannot modify subject (COM): {comEx.Message}");
                    }
                }

                mail.Save();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error applying classification: {ex.Message}");
                throw; // Re-throw to handle in calling method
            }
        }

        private void CreateCategoriesIfNeeded(Outlook._NameSpace session)
        {
            try
            {
                var categories = session.Categories;

                foreach (var kv in config.ClassificationSettings.Classifications)
                {
                    var cfg = kv.Value;
                    if (cfg == null) continue;

                    string name = !string.IsNullOrEmpty(cfg.CategoryPrefix) ? cfg.CategoryPrefix : kv.Key;

                    Outlook.OlCategoryColor color = ParseCategoryColor(cfg.CategoryColor);

                    CreateCategoryIfNotExists(categories, name, color);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating categories: {ex.Message}");
            }
        }

        private void CreateCategoryIfNotExists(Outlook.Categories categories, string name, Outlook.OlCategoryColor color)
        {
            try
            {
                bool exists = false;
                foreach (Outlook.Category cat in categories)
                {
                    if (cat.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    {
                        exists = true;
                        break;
                    }
                }

                if (!exists)
                {
                    categories.Add(name, color);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating category {name}: {ex.Message}");
            }
        }

        private void ShowProgressForm(int totalEmails)
        {
            // Close any existing progress form
            CloseProgressForm();

            // Create new progress form on UI thread
            if (System.Windows.Forms.Application.MessageLoop)
            {
                _progressForm = new ProgressForm(config, FormatTemplate, totalEmails);
                _progressForm.Show();
            }
            else
            {
                System.Windows.Forms.Application.Run(new ProgressForm(config, FormatTemplate, totalEmails));
            }
        }

        private void UpdateProgress(string status, int current, int total)
        {
            if (_progressForm == null || _progressForm.IsDisposed) return;

            if (_progressForm.InvokeRequired)
            {
                _progressForm.Invoke(new Action(() => UpdateProgress(status, current, total)));
                return;
            }

            try
            {
                _progressForm.UpdateProgress(status, current, total);
            }
            catch (ObjectDisposedException)
            {
                // Form was disposed, ignore
            }
        }

        private void CloseProgressForm()
        {
            if (_progressForm == null) return;

            if (_progressForm.InvokeRequired)
            {
                _progressForm.Invoke(new Action(() => CloseProgressForm()));
                return;
            }

            try
            {
                if (!_progressForm.IsDisposed)
                {
                    _progressForm.Close();
                    _progressForm.Dispose();
                }
            }
            catch (ObjectDisposedException)
            {
                // Already disposed, ignore
            }
            finally
            {
                _progressForm = null;
            }
        }

        private Outlook.OlImportance ParseImportance(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return Outlook.OlImportance.olImportanceNormal;

            // Try exact enum name
            Outlook.OlImportance parsed;
            if (Enum.TryParse<Outlook.OlImportance>(s, true, out parsed))
                return parsed;

            // Common synonyms
            string t = s.Trim().ToLowerInvariant();
            if (t == "high") return Outlook.OlImportance.olImportanceHigh;
            if (t == "low") return Outlook.OlImportance.olImportanceLow;
            if (t == "normal" || t == "medium") return Outlook.OlImportance.olImportanceNormal;

            return Outlook.OlImportance.olImportanceNormal;
        }

        private bool TryParseFlagIcon(string s, out Outlook.OlFlagIcon icon)
        {
            // Default
            icon = Outlook.OlFlagIcon.olNoFlagIcon;
            if (string.IsNullOrWhiteSpace(s)) return false;

            // Exact enum name?
            Outlook.OlFlagIcon parsed;
            if (Enum.TryParse<Outlook.OlFlagIcon>(s, true, out parsed))
            {
                icon = parsed;
                return parsed != Outlook.OlFlagIcon.olNoFlagIcon;
            }

            // Friendly names → Outlook enum
            string t = s.Trim().ToLowerInvariant();
            if (t == "red") { icon = Outlook.OlFlagIcon.olRedFlagIcon; return true; }
            if (t == "yellow") { icon = Outlook.OlFlagIcon.olYellowFlagIcon; return true; }
            if (t == "blue") { icon = Outlook.OlFlagIcon.olBlueFlagIcon; return true; }
            if (t == "green") { icon = Outlook.OlFlagIcon.olGreenFlagIcon; return true; }
            if (t == "orange") { icon = Outlook.OlFlagIcon.olOrangeFlagIcon; return true; }
            if (t == "purple") { icon = Outlook.OlFlagIcon.olPurpleFlagIcon; return true; }
            if (t == "no" || t == "none" || t == "off") { icon = Outlook.OlFlagIcon.olNoFlagIcon; return false; }

            return false;
        }

        private Outlook.OlCategoryColor ParseCategoryColor(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return Outlook.OlCategoryColor.olCategoryColorNone;

            // Exact enum name first
            Outlook.OlCategoryColor parsed;
            if (Enum.TryParse<Outlook.OlCategoryColor>(s, true, out parsed))
                return parsed;

            // Friendly basic colors
            string t = s.Trim().ToLowerInvariant();
            if (t == "red") return Outlook.OlCategoryColor.olCategoryColorRed;
            if (t == "blue") return Outlook.OlCategoryColor.olCategoryColorBlue;
            if (t == "yellow") return Outlook.OlCategoryColor.olCategoryColorYellow;
            if (t == "green") return Outlook.OlCategoryColor.olCategoryColorGreen;
            if (t == "orange") return Outlook.OlCategoryColor.olCategoryColorOrange;
            if (t == "purple") return Outlook.OlCategoryColor.olCategoryColorPurple;
            if (t == "teal") return Outlook.OlCategoryColor.olCategoryColorTeal;
            if (t == "gray" || t == "grey") return Outlook.OlCategoryColor.olCategoryColorDarkGray;

            return Outlook.OlCategoryColor.olCategoryColorNone;
        }

        // Add the ProgressForm class at the end of your Ribbon1 class file
        public class ProgressForm : Form
        {
            private Label _statusLabel;
            private ProgressBar _progressBar;
            private Label _percentLabel;
            private AppConfig _config;
            private Func<string, Dictionary<string, string>, string> _formatTemplate;

            public ProgressForm(AppConfig config, Func<string, Dictionary<string, string>, string> formatTemplate, int totalEmails)
            {
                _config = config;
                _formatTemplate = formatTemplate;

                int width = config.WindowSettings != null && config.WindowSettings.MainWindow != null
                            ? config.WindowSettings.MainWindow.Width
                            : 520;
                int height = config.WindowSettings != null && config.WindowSettings.MainWindow != null
                            ? Math.Max(180, config.WindowSettings.MainWindow.Height / 3)
                            : 220;

                this.Width = width;
                this.Height = height;
                this.Text = config.Messages.AnalyzingEmail;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.BackColor = ConfigManager.ParseColor(config.UiSettings.Colors.DarkBackground);
                this.ForeColor = ConfigManager.ParseColor(config.UiSettings.Colors.TextColor);
                this.Font = ConfigManager.ParseFont(config.UiSettings.Fonts.Default);

                _statusLabel = new Label
                {
                    Name = "StatusLabel",
                    Text = config.Messages.Loading,
                    ForeColor = ConfigManager.ParseColor(config.UiSettings.Colors.TextColor),
                    BackColor = Color.Transparent,
                    AutoSize = false,
                    Size = new Size(width - 40, 60),
                    Location = new Point(20, 20),
                    TextAlign = ContentAlignment.MiddleLeft
                };

                _progressBar = new ProgressBar
                {
                    Name = "ProgressBar",
                    Size = new Size(width - 40, 22),
                    Location = new Point(20, 90),
                    Maximum = Math.Max(0, totalEmails),
                    Value = 0,
                    Style = ProgressBarStyle.Continuous
                };

                _percentLabel = new Label
                {
                    Name = "PercentLabel",
                    Text = "0%",
                    ForeColor = ConfigManager.ParseColor(config.UiSettings.Colors.TextColor),
                    BackColor = Color.Transparent,
                    AutoSize = false,
                    Size = new Size(width - 40, 22),
                    Location = new Point(20, 120),
                    TextAlign = ContentAlignment.MiddleCenter
                };

                this.Controls.Add(_statusLabel);
                this.Controls.Add(_progressBar);
                this.Controls.Add(_percentLabel);
            }

            public void UpdateProgress(string status, int current, int total)
            {
                if (this.IsDisposed) return;

                if (_statusLabel != null) _statusLabel.Text = status;
                if (_progressBar != null && total > 0)
                {
                    _progressBar.Maximum = Math.Max(total, 0);
                    _progressBar.Value = Math.Min(Math.Max(current, 0), Math.Max(total, 0));
                }
                if (_percentLabel != null && total > 0)
                {
                    int percentage = total > 0 ? (int)((double)current / total * 100) : 0;
                    _percentLabel.Text = _formatTemplate(_config.ClassificationSettings.Progress.PercentageTemplate,
                        new Dictionary<string, string>
                        {
                    {"{PERCENTAGE}", percentage.ToString()},
                    {"{CURRENT}", current.ToString()},
                    {"{TOTAL}", total.ToString()}
                        });
                }
            }
        }

        
    }
}

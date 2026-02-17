# SmartMailAssistant

A smart Outlook VSTO Add-in that leverages AI to enhance email productivity. This tool integrates directly into the Outlook Ribbon to provide summarization, translation, smart replies, and automated email classification using configurable AI models.

## Features

* Summarize: Instantly generates a bullet-point summary of the selected email, stripping away signatures and clutter.
* Translate: Translates email content into your preferred language with support for Right-to-Left (RTL) languages like Arabic, Hebrew, and Urdu.
* Smart Reply: Analyzes incoming emails and generates context-aware, professional responses based on configurable protocols.
* AI & Rule-Based Classification: Automatically categorizes unread emails (e.g., Urgent, Work, Social) using a hybrid system of keyword rules and AI analysis.
* Customizable UI: Fully configurable fonts, colors, and loading states via a configuration file.
* Rate Limiting & Retry: Built-in exponential backoff and rate limiting to handle API constraints gracefully.

## Tech Stack

* Framework: .NET Framework 4.8 


* Platform: Microsoft Outlook VSTO Add-in 


* Language: C# 


* Dependencies: Newtonsoft.Json 


* UI: Windows Forms (integrated into Outlook Ribbon) 



## Configuration

The application is driven by a config.txt file located in the application directory. You must configure your API credentials before use.

Example Configuration Structure:

ApiSettings:
ApiKey: "your_api_key_here"
ApiEndpoint: "your_api_endpoint"
ModelName: "your_model_name"
TimeoutSeconds: 60

LanguageSettings:
DefaultTranslationLanguage: "Arabic"
FallbackLanguage: "English"

ClassificationSettings:
PreProcessingRules:
EnableRuleBasedClassification: true
Rules:
- Name: "Urgent Rule"
Classification: "URGENT"
Conditions:
SubjectKeywords: ["ASAP", "Deadline"]

## Installation & Setup

1. Clone the repository
git clone [https://github.com/yourusername/SmartMailAssistant.git]()
2. Open in Visual Studio
Open SmartMailAssistant1.csproj in Visual Studio (ensure Office/SharePoint development workload is installed).
3. Restore Dependencies
The project uses NuGet packages. Restore them to ensure Newtonsoft.Json and others are available.
4. Configure config.txt
* Locate config.txt in the project root.
* Update ApiKey, ApiEndpoint, and ModelName with your AI provider details.
* Set "Copy to Output Directory" to "Copy always" or ensure it exists in the build output folder.


5. Build and Run
Press F5 to build and start Outlook with the Add-in attached.

## Usage

1. Ribbon Controls: Navigate to the SmartMailAssistant tab in the Outlook Ribbon.


2. Processing Emails: Select an email and click Summarize, Translate, or Suggest Reply. A popup window will stream the AI response.
3. Classifying: Click Classify Emails to scan unread items in the current folder. The add-in will apply Categories and Flags (e.g., Red Flag for Urgent) based on the content.


## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

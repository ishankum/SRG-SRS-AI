using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace SrsAi.Functions
{
    public class ClassificationResult
    {
        public string ProjectName { get; set; } = "RSCS";
        public string MatchedSource { get; set; } = "DefaultFallback";
        public double Confidence { get; set; } = 0.5;
        public string Details { get; set; } = "";
    }

    public class ProjectClassifier
    {
        private readonly IConfiguration _configuration;
        private readonly ILogger _logger;

        // Project aliases and domain identifiers
        private static readonly Dictionary<string, string[]> ProjectAliases = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
        {
            {
                "RSCS", new[] { "rscs", "rscslive", "rscs_main", "rscs main", "rscs_client", "rscs live", "rscs_mainlive" }
            },
            {
                "CBRE", new[] { "cbre", "cbre_main", "cbre main", "cbre_client", "cbrelive", "cbre live" }
            },
            {
                "AstuteDFM", new[] { "astutedfm", "astute dfm", "astutedfm_main", "astute", "dfm" }
            }
        };

        public ProjectClassifier(IConfiguration configuration, ILogger logger)
        {
            _configuration = configuration;
            _logger = logger;
        }

        public ClassificationResult Classify(string transcriptRawJson)
        {
            string defaultProject = _configuration["DefaultProject"] ?? "RSCS";
            string activeProjectsStr = _configuration["ActiveProjects"] ?? "RSCS,CBRE,AstuteDFM";
            var activeProjects = activeProjectsStr
                .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(p => p.Trim().ToUpper())
                .ToList();

            if (string.IsNullOrWhiteSpace(transcriptRawJson))
            {
                _logger.LogWarning($"Transcript payload is empty. Defaulting project to '{defaultProject}'.");
                return new ClassificationResult { ProjectName = defaultProject, MatchedSource = "EmptyTranscriptFallback", Confidence = 0.5 };
            }

            string title = "";
            string overview = "";
            string keywords = "";
            string fullSpokenText = "";
            List<string> speakerEmails = new List<string>();

            try
            {
                using var doc = JsonDocument.Parse(transcriptRawJson);
                var root = doc.RootElement;

                // Fireflies GraphQL wrapper: { "data": { "transcript": { ... } } } or direct { "transcript": { ... } }
                JsonElement transcriptElem = root;
                if (root.TryGetProperty("data", out var dataElem) && dataElem.TryGetProperty("transcript", out var tElem1))
                {
                    transcriptElem = tElem1;
                }
                else if (root.TryGetProperty("transcript", out var tElem2))
                {
                    transcriptElem = tElem2;
                }

                if (transcriptElem.ValueKind == JsonValueKind.Object)
                {
                    if (transcriptElem.TryGetProperty("title", out var titleProp))
                        title = titleProp.GetString() ?? "";

                    if (transcriptElem.TryGetProperty("summary", out var summaryProp))
                    {
                        if (summaryProp.TryGetProperty("overview", out var ovProp))
                            overview = ovProp.GetString() ?? "";

                        if (summaryProp.TryGetProperty("keywords", out var kwProp))
                        {
                            if (kwProp.ValueKind == JsonValueKind.Array)
                            {
                                keywords = string.Join(" ", kwProp.EnumerateArray().Select(k => k.GetString()));
                            }
                            else if (kwProp.ValueKind == JsonValueKind.String)
                            {
                                keywords = kwProp.GetString() ?? "";
                            }
                        }
                    }

                    if (transcriptElem.TryGetProperty("sentences", out var sentencesProp) && sentencesProp.ValueKind == JsonValueKind.Array)
                    {
                        var sb = new System.Text.StringBuilder();
                        foreach (var s in sentencesProp.EnumerateArray())
                        {
                            if (s.TryGetProperty("text", out var textProp))
                            {
                                sb.Append(textProp.GetString()).Append(" ");
                            }
                        }
                        fullSpokenText = sb.ToString();
                    }
                    else if (transcriptElem.TryGetProperty("text", out var rawTextProp))
                    {
                        fullSpokenText = rawTextProp.GetString() ?? "";
                    }

                    if (transcriptElem.TryGetProperty("speakers", out var speakersProp) && speakersProp.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var sp in speakersProp.EnumerateArray())
                        {
                            if (sp.TryGetProperty("email", out var emailProp))
                            {
                                string e = emailProp.GetString() ?? "";
                                if (!string.IsNullOrWhiteSpace(e)) speakerEmails.Add(e.ToLower());
                            }
                        }
                    }
                }
            }
            catch (Exception parseEx)
            {
                _logger.LogWarning($"Could not parse structured Fireflies JSON for classification: {parseEx.Message}. Performing raw string scan.");
                fullSpokenText = transcriptRawJson;
            }

            // TIER 1: Check Meeting Title (Strongest Signal)
            if (!string.IsNullOrWhiteSpace(title))
            {
                foreach (var project in activeProjects)
                {
                    if (MatchesProject(title, project))
                    {
                        _logger.LogInformation($"[ProjectClassifier] Tier 1 Match: Title '{title}' matched project '{project}'.");
                        return new ClassificationResult
                        {
                            ProjectName = project,
                            MatchedSource = "MeetingTitle",
                            Confidence = 0.98,
                            Details = $"Matched in meeting title: \"{title}\""
                        };
                    }
                }
            }

            // TIER 2: Check Summary Overview & Keywords
            string summaryCombined = $"{overview} {keywords}";
            if (!string.IsNullOrWhiteSpace(summaryCombined))
            {
                foreach (var project in activeProjects)
                {
                    if (MatchesProject(summaryCombined, project))
                    {
                        _logger.LogInformation($"[ProjectClassifier] Tier 2 Match: Summary/Keywords matched project '{project}'.");
                        return new ClassificationResult
                        {
                            ProjectName = project,
                            MatchedSource = "SummaryAndKeywords",
                            Confidence = 0.88,
                            Details = $"Matched in summary keywords/overview"
                        };
                    }
                }
            }

            // TIER 3: Check Speaker Email Domains
            foreach (var email in speakerEmails)
            {
                if (email.Contains("cbre.com") && activeProjects.Contains("CBRE"))
                {
                    _logger.LogInformation($"[ProjectClassifier] Tier 3 Match: Speaker email domain '{email}' matched CBRE.");
                    return new ClassificationResult
                    {
                        ProjectName = "CBRE",
                        MatchedSource = "SpeakerEmailDomain",
                        Confidence = 0.85,
                        Details = $"Matched attendee email '{email}'"
                    };
                }
            }

            // TIER 4: Explicit Spoken Mentions Frequency in Transcript Body
            var mentionCounts = new Dictionary<string, int>();
            foreach (var project in activeProjects)
            {
                int count = CountMentions(fullSpokenText, project);
                if (count > 0)
                {
                    mentionCounts[project] = count;
                }
            }

            if (mentionCounts.Count > 0)
            {
                var topMatch = mentionCounts.OrderByDescending(kv => kv.Value).First();
                _logger.LogInformation($"[ProjectClassifier] Tier 4 Match: Highest transcript mention count is project '{topMatch.Key}' with {topMatch.Value} occurrences.");
                return new ClassificationResult
                {
                    ProjectName = topMatch.Key,
                    MatchedSource = "TranscriptMentionFrequency",
                    Confidence = 0.75,
                    Details = $"Mentioned {topMatch.Value} times in conversation"
                };
            }

            // TIER 5: Fallback to Default Project
            _logger.LogWarning($"[ProjectClassifier] Inconclusive project detection. Defaulting to '{defaultProject}'.");
            return new ClassificationResult
            {
                ProjectName = defaultProject,
                MatchedSource = "DefaultFallback",
                Confidence = 0.50,
                Details = $"No explicit project markers found. Fallback to default."
            };
        }

        private static bool MatchesProject(string text, string projectName)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;

            // Direct word boundary match for ProjectName (e.g., \bCBRE\b, \bRSCS\b)
            if (Regex.IsMatch(text, $@"\b{Regex.Escape(projectName)}\b", RegexOptions.IgnoreCase))
            {
                return true;
            }

            // Check aliases
            if (ProjectAliases.TryGetValue(projectName, out var aliases))
            {
                foreach (var alias in aliases)
                {
                    if (Regex.IsMatch(text, $@"\b{Regex.Escape(alias)}\b", RegexOptions.IgnoreCase))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static int CountMentions(string text, string projectName)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;

            int count = Regex.Matches(text, $@"\b{Regex.Escape(projectName)}\b", RegexOptions.IgnoreCase).Count;

            if (ProjectAliases.TryGetValue(projectName, out var aliases))
            {
                foreach (var alias in aliases)
                {
                    count += Regex.Matches(text, $@"\b{Regex.Escape(alias)}\b", RegexOptions.IgnoreCase).Count;
                }
            }

            return count;
        }
    }
}

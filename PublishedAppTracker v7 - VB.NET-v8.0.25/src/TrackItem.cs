using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace PublishedAppTracker
{
    public class TrackItem
    {
        // Core fields
        public string ProgramName { get; set; } = "";
        public string TrackURL { get; set; } = "";
        public string StartString { get; set; } = "";
        public string StopString { get; set; } = "";
        public string DownloadURL { get; set; } = "";
        public string Version { get; set; } = "";
        public string FilePath { get; set; } = "";

        // Tracking fields
        public string LatestVersion { get; set; } = "";
        public string TrackBlockHash { get; set; } = "";
        public string PreviousHash { get; set; } = "";
        public string LastChecked { get; set; } = "";
        public string DownloadSizeKb { get; set; } = "";
        public string TrackStatus { get; set; } = "unchecked";

        // Metadata fields
        public string ReleaseDate { get; set; } = "";
        public string CreationDate { get; set; } = "";
        public string ModificationDate { get; set; } = "";
        public string PublisherName { get; set; } = "";
        public string SuiteName { get; set; } = "";

        // UI fields
        public bool IsSelected { get; set; } = false;

		// Add these alongside existing properties
		public string TrackURLStatus { get; set; } = "";
		public string StartStringStatus { get; set; } = "";
		public string StopStringStatus { get; set; } = "";
		public string HashStatus { get; set; } = "";
		public string LatestVersionStatus { get; set; } = "";
		public string DownloadURLStatus { get; set; } = "";
		public string DownloadSizeStatus { get; set; } = "";
		
        // ============================
        // Track Block Extraction
        // ============================

        public static string ExtractTrackBlock(string source, string startString, string stopString)
        {
            if (string.IsNullOrEmpty(source) ||
                string.IsNullOrEmpty(startString) ||
                string.IsNullOrEmpty(stopString))
                return null;

            bool startIsRegex = IsRegexPattern(startString);
            bool stopIsRegex = IsRegexPattern(stopString);

            // Normalize line endings
            string normalizedSource = source.Replace("\r\n", "\n").Replace("\r", "\n");

            int startIdx = -1;
            int startMatchLen = 0;
            int stopIdx = -1;
            int stopMatchLen = 0;

            // Find start string
            if (startIsRegex)
            {
                var match = FindRegexMatch(normalizedSource, startString, 0);
                if (match != null)
                {
                    startIdx = match.Index;
                    startMatchLen = match.Length;
                }
            }
            else
            {
                string normalizedStart = startString.Replace("\r\n", "\n").Replace("\r", "\n");

                // Exact match first
                startIdx = normalizedSource.IndexOf(normalizedStart, StringComparison.Ordinal);
                startMatchLen = normalizedStart.Length;

                // Whitespace-collapsed match fallback
                if (startIdx < 0)
                {
                    string collapsedSource = CollapseWhitespace(normalizedSource);
                    string collapsedStart = CollapseWhitespace(normalizedStart);
                    int collapsedIdx = collapsedSource.IndexOf(collapsedStart, StringComparison.Ordinal);

                    if (collapsedIdx >= 0)
                    {
                        startIdx = MapCollapsedToOriginal(normalizedSource, collapsedIdx);
                        int startEnd = MapCollapsedToOriginal(normalizedSource,
                            collapsedIdx + collapsedStart.Length);
                        startMatchLen = startEnd - startIdx;
                    }
                }
            }

            if (startIdx < 0)
                return null;

            // Find stop string (searching after the start match)
            int searchAfter = startIdx + startMatchLen;

            if (stopIsRegex)
            {
                var match = FindRegexMatch(normalizedSource, stopString, searchAfter);
                if (match != null)
                {
                    stopIdx = match.Index;
                    stopMatchLen = match.Length;
                }
            }
            else
            {
                string normalizedStop = stopString.Replace("\r\n", "\n").Replace("\r", "\n");

                stopIdx = normalizedSource.IndexOf(normalizedStop, searchAfter, StringComparison.Ordinal);
                stopMatchLen = normalizedStop.Length;

                if (stopIdx < 0)
                {
                    string collapsedSource = CollapseWhitespace(normalizedSource);
                    string collapsedStop = CollapseWhitespace(normalizedStop);

                    // Need to find collapsed position corresponding to searchAfter
                    string collapsedAfterStart = CollapseWhitespace(normalizedSource.Substring(0, searchAfter));
                    int collapsedSearchAfter = collapsedAfterStart.Length;

                    int collapsedIdx = collapsedSource.IndexOf(collapsedStop, collapsedSearchAfter,
                        StringComparison.Ordinal);

                    if (collapsedIdx >= 0)
                    {
                        stopIdx = MapCollapsedToOriginal(normalizedSource, collapsedIdx);
                        int stopEnd = MapCollapsedToOriginal(normalizedSource,
                            collapsedIdx + collapsedStop.Length);
                        stopMatchLen = stopEnd - stopIdx;
                    }
                }
            }

            if (stopIdx < 0)
                return null;

            int blockEnd = stopIdx + stopMatchLen;
            return normalizedSource.Substring(startIdx, blockEnd - startIdx);
        }

        /// <summary>
        /// Determines if a string contains regex special characters that indicate
        /// it should be treated as a regex pattern rather than a literal string.
        /// </summary>
        public static bool IsRegexPattern(string input)
        {
            if (string.IsNullOrEmpty(input))
                return false;

            // Check for common regex patterns
            // We look for patterns that are unlikely to appear in literal HTML
            string[] regexIndicators = { ".*", ".+", "\\d", "\\w", "\\s", "[^", "(?:",
                                          "(?=", "(?!", "\\b", "{\\d", ".{" };

            foreach (string indicator in regexIndicators)
            {
                if (input.Contains(indicator))
                    return true;
            }

            // Also check for character classes like [a-z] but not HTML attributes like [class=...]
            // A regex character class typically has a hyphen between single characters: [a-z], [0-9]
            for (int i = 0; i < input.Length - 4; i++)
            {
                if (input[i] == '[' && i + 4 < input.Length)
                {
                    // Check if it looks like [x-y] pattern (regex range)
                    if (input[i + 2] == '-' && input[i + 4] == ']')
                        return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Finds a regex match in the source string starting from the given position.
        /// The pattern has line endings normalized and is matched with singleline mode
        /// so that . matches newlines too.
        /// </summary>
        /// <summary>
        /// Finds a regex match in the source string starting from the given position.
        /// The pattern has line endings normalized. Greedy quantifiers (.*  .+) are
        /// automatically converted to non-greedy (.*?  .+?) so they don't overshoot
        /// past the intended boundary in HTML content.
        /// </summary>
        public static System.Text.RegularExpressions.Match FindRegexMatch(
            string source, string pattern, int startAt)
        {
            try
            {
                string normalizedPattern = pattern.Replace("\r\n", "\n").Replace("\r", "\n");

                // Convert greedy quantifiers to non-greedy to prevent overshooting.
                // .*  becomes .*?    .+  becomes .+?
                // But don't double-convert if already non-greedy (.*? stays .*?)
                normalizedPattern = System.Text.RegularExpressions.Regex.Replace(
                    normalizedPattern, @"\.\*(?!\?)", ".*?");
                normalizedPattern = System.Text.RegularExpressions.Regex.Replace(
                    normalizedPattern, @"\.\+(?!\?)", ".+?");

                var regex = new System.Text.RegularExpressions.Regex(
                    normalizedPattern,
                    System.Text.RegularExpressions.RegexOptions.Singleline |
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase |
                    System.Text.RegularExpressions.RegexOptions.Compiled);

                var match = regex.Match(source, startAt);
                if (match.Success)
                    return match;
                return null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Collapses all runs of whitespace (spaces, tabs, newlines) into a single space.
        /// This allows matching HTML content regardless of indentation differences.
        /// </summary>
        private static string CollapseWhitespace(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            StringBuilder sb = new StringBuilder(text.Length);
            bool lastWasWhitespace = false;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == ' ' || c == '\t' || c == '\n' || c == '\r')
                {
                    if (!lastWasWhitespace)
                    {
                        sb.Append(' ');
                        lastWasWhitespace = true;
                    }
                    // Skip additional whitespace characters
                }
                else
                {
                    sb.Append(c);
                    lastWasWhitespace = false;
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// Maps a character position in the collapsed string back to the
        /// corresponding position in the original string.
        /// </summary>
        private static int MapCollapsedToOriginal(string original, int collapsedPos)
        {
            int collapsed = 0;
            bool lastWasWhitespace = false;

            for (int i = 0; i < original.Length; i++)
            {
                if (collapsed == collapsedPos)
                    return i;

                char c = original[i];
                if (c == ' ' || c == '\t' || c == '\n' || c == '\r')
                {
                    if (!lastWasWhitespace)
                    {
                        collapsed++;
                        lastWasWhitespace = true;
                    }
                    // Additional whitespace chars don't advance collapsed position
                }
                else
                {
                    collapsed++;
                    lastWasWhitespace = false;
                }
            }

            // If we reached the end and collapsed matches, return end of string
            if (collapsed == collapsedPos)
                return original.Length;

            return -1;
        }

        public static string ExtractVersion(string trackBlock)
        {
            if (string.IsNullOrEmpty(trackBlock))
                return "";

            // Strategy 1: Look for common version prefixes in the track block
            string[] prefixes = { "v", "V", "ver ", "Ver ", "version ", "Version ",
                                  "release ", "Release " };

            foreach (string prefix in prefixes)
            {
                int idx = trackBlock.IndexOf(prefix, StringComparison.Ordinal);
                if (idx >= 0)
                {
                    string afterPrefix = trackBlock.Substring(idx + prefix.Length).Trim();
                    string ver = ExtractVersionFromStart(afterPrefix);
                    if (!string.IsNullOrEmpty(ver))
                        return ver;
                }
            }

            // Strategy 2: Find first version-like pattern anywhere
            return ExtractFirstVersionPattern(trackBlock);
        }

		private static string ExtractVersionFromStart(string text)
		{
		    StringBuilder version = new StringBuilder();

		    for (int i = 0; i < text.Length; i++)
		    {
		        char c = text[i];
		        if (char.IsDigit(c) || c == '.' || c == '-')
		        {
		            version.Append(c);
		        }
		        else
		        {
		            break;
		        }
		    }

		    string candidate = version.ToString().Trim('.').Trim('-');

		    // Normalize: if it uses hyphens as separators (e.g., 10-1-25), convert to dots
		    if (!candidate.Contains(".") && candidate.Contains("-"))
		    {
		        candidate = candidate.Replace('-', '.');
		    }

		    if (candidate.Contains(".") && candidate.Length >= 3)
		    {
		        return candidate;
		    }

		    return "";
		}

		private static string ExtractFirstVersionPattern(string text)
		{
		    StringBuilder version = new StringBuilder();
		    bool inVersion = false;

		    for (int i = 0; i < text.Length; i++)
		    {
		        char c = text[i];

		        if (!inVersion)
		        {
		            if (char.IsDigit(c))
		            {
		                inVersion = true;
		                version.Append(c);
		            }
		        }
		        else
		        {
		            if (char.IsDigit(c) || c == '.' || c == '-')
		            {
		                version.Append(c);
		            }
		            else
		            {
		                string candidate = NormalizeVersionCandidate(version.ToString());
		                if (candidate != null)
		                    return candidate;

		                version.Clear();
		                inVersion = false;
		            }
		        }
		    }

		    if (inVersion)
		    {
		        string candidate = NormalizeVersionCandidate(version.ToString());
		        if (candidate != null)
		            return candidate;
		    }

		    return "";
		}

		/// <summary>
		/// Cleans up a raw version candidate string. Converts hyphen-separated
		/// versions to dot-separated (e.g., "10-1-25" → "10.1.25") and validates
		/// that it looks like a real version number (contains a separator and
		/// has at least 3 characters).
		/// Returns the normalized version string, or null if invalid.
		/// </summary>
		private static string NormalizeVersionCandidate(string raw)
		{
		    string candidate = raw.Trim('.').Trim('-');

		    if (string.IsNullOrEmpty(candidate))
		        return null;

		    // If it uses hyphens but no dots, treat hyphens as version separators
		    if (!candidate.Contains(".") && candidate.Contains("-"))
		    {
		        candidate = candidate.Replace('-', '.');
		    }

		    // Also handle mixed: strip trailing hyphens/dots after conversion
		    candidate = candidate.Trim('.');

		    if (candidate.Contains(".") && candidate.Length >= 3)
		        return candidate;

		    return null;
		}

        public static string ComputeHash(string input)
        {
            if (string.IsNullOrEmpty(input))
                return "";

            using (SHA256 sha = SHA256.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(input);
                byte[] hash = sha.ComputeHash(bytes);

                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < 8; i++)
                {
                    sb.Append(hash[i].ToString("x2"));
                }
                return sb.ToString();
            }
        }

        // ============================
        // Apply Check Result
        // ============================

        public CheckResult ApplyCheck(string pageSource, long downloadBytes)
        {
            CheckResult result = new CheckResult();
            result.ProgramName = ProgramName;

            // URL worked if we got here
            TrackURLStatus = "ok";

            // Handle download size
            if (downloadBytes > 0)
            {
                long localSize = 0;
                long.TryParse(DownloadSizeKb, out localSize);
                long remoteKb = downloadBytes / 1024;

                if (localSize > 0 && Math.Abs(localSize - remoteKb) <= (remoteKb * 0.1))
                    DownloadSizeStatus = "ok";
                else if (localSize > 0)
                    DownloadSizeStatus = "changed";
                else
                    DownloadSizeStatus = "ok";

                DownloadSizeKb = remoteKb.ToString();
            }

            LastChecked = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            ModificationDate = LastChecked;

            string trackBlock = ExtractTrackBlock(pageSource, StartString, StopString);
            bool hasTrackStrings = !string.IsNullOrEmpty(StartString) &&
                                   !string.IsNullOrEmpty(StopString);

            if (trackBlock == null)
            {
                if (!hasTrackStrings)
                {
                    trackBlock = pageSource;
                    result.Note = "No start/stop strings - hashing full page";
                    StartStringStatus = "";
                    StopStringStatus = "";
                }
                else
                {
                    TrackStatus = "error";
                    result.Status = "error";
                    result.Note = "Track block not found on page";

                    // Figure out which string failed
                    string normalizedSource = pageSource.Replace("\r\n", "\n").Replace("\r", "\n");
                    string normalizedStart = StartString.Replace("\r\n", "\n").Replace("\r", "\n");
                    int startIdx = normalizedSource.IndexOf(normalizedStart, StringComparison.Ordinal);
                    StartStringStatus = startIdx >= 0 ? "ok" : "error";
                    StopStringStatus = "error";
                    HashStatus = "";
                    return result;
                }
            }
            else
            {
                StartStringStatus = "ok";
                StopStringStatus = "ok";
            }

            // Auto-extract version from track block
            // Only extract version if we have start/stop strings defining a specific block
            if (hasTrackStrings)
            {
                string extractedVersion = ExtractVersion(trackBlock);
                if (!string.IsNullOrEmpty(extractedVersion))
                {
                    LatestVersion = extractedVersion;
                    result.ExtractedVersion = extractedVersion;
                    LatestVersionStatus = (extractedVersion == Version) ? "ok" : "changed";

                    // If Version is empty, auto-fill it on first check
                    if (string.IsNullOrEmpty(Version))
                    {
                        Version = extractedVersion;
                    }
                }
            }
            else
            {
                LatestVersion = "";
                LatestVersionStatus = "";
            }

            // Compute hash
            string newHash = ComputeHash(trackBlock);

            if (string.IsNullOrEmpty(TrackBlockHash))
            {
                PreviousHash = "";
                TrackBlockHash = newHash;
                TrackStatus = "new";
                result.Status = "new";
                result.Note = "First check — version: " + result.ExtractedVersion;
                HashStatus = "ok";
            }
            else if (newHash == TrackBlockHash)
            {
                TrackStatus = "unchanged";
                result.Status = "unchanged";
                result.Note = "No change detected";
                HashStatus = "ok";
            }
            else
            {
                PreviousHash = TrackBlockHash;
                TrackBlockHash = newHash;
                TrackStatus = "changed";
                result.Status = "changed";
                result.Note = "CHANGE DETECTED! Version: " + result.ExtractedVersion;
                HashStatus = "changed";
            }

            result.Hash = newHash;
            result.TrackBlock = trackBlock;
            return result;
        }

        // ============================
        // Save (XML)
        // ============================

        public void SaveToFile(string path = null)
        {
            if (path != null)
                FilePath = path;

            if (string.IsNullOrEmpty(FilePath))
                throw new InvalidOperationException("No file path set for track item.");

            if (string.IsNullOrEmpty(CreationDate))
                CreationDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

            ModificationDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.IndentChars = "  ";
            settings.Encoding = Encoding.UTF8;

            using (XmlWriter writer = XmlWriter.Create(FilePath, settings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("AppTrack");

                writer.WriteStartElement("Track");
                writer.WriteElementString("ProgramName", ProgramName ?? "");
                writer.WriteElementString("TrackURL", TrackURL ?? "");
                writer.WriteElementString("StartString", StartString ?? "");
                writer.WriteElementString("StopString", StopString ?? "");
                writer.WriteElementString("DownloadURL", DownloadURL ?? "");
                writer.WriteElementString("Version", Version ?? "");
                writer.WriteElementString("LatestVersion", LatestVersion ?? "");
                writer.WriteElementString("LastChecked", LastChecked ?? "");
                writer.WriteElementString("TrackBlockHash", TrackBlockHash ?? "");
                writer.WriteElementString("PreviousHash", PreviousHash ?? "");
                writer.WriteElementString("DownloadSizeKb", DownloadSizeKb ?? "");
                writer.WriteElementString("TrackStatus", TrackStatus ?? "unchecked");
                writer.WriteElementString("ReleaseDate", ReleaseDate ?? "");
                writer.WriteElementString("CreationDate", CreationDate ?? "");
                writer.WriteElementString("ModificationDate", ModificationDate ?? "");
                writer.WriteElementString("PublisherName", PublisherName ?? "");
                writer.WriteElementString("SuiteName", SuiteName ?? "");
                writer.WriteEndElement();

                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        // ============================
        // Load (XML)
        // ============================

        public static TrackItem LoadFromFile(string path)
        {
            TrackItem item = new TrackItem();
            item.FilePath = path;

            if (!File.Exists(path))
                return item;

            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(path);

                XmlNode track = doc.SelectSingleNode("//Track");
                if (track == null)
                    return item;

                item.ProgramName = GetNodeText(track, "ProgramName");
                item.TrackURL = GetNodeText(track, "TrackURL");
                item.StartString = GetNodeText(track, "StartString");
                item.StopString = GetNodeText(track, "StopString");
                item.DownloadURL = GetNodeText(track, "DownloadURL");
                item.Version = GetNodeText(track, "Version");
                item.LatestVersion = GetNodeText(track, "LatestVersion");
                item.LastChecked = GetNodeText(track, "LastChecked");
                item.TrackBlockHash = GetNodeText(track, "TrackBlockHash");
                item.PreviousHash = GetNodeText(track, "PreviousHash");
                item.DownloadSizeKb = GetNodeText(track, "DownloadSizeKb");
                item.TrackStatus = GetNodeText(track, "TrackStatus");
                item.ReleaseDate = GetNodeText(track, "ReleaseDate");
                item.CreationDate = GetNodeText(track, "CreationDate");
                item.ModificationDate = GetNodeText(track, "ModificationDate");
                item.PublisherName = GetNodeText(track, "PublisherName");
                item.SuiteName = GetNodeText(track, "SuiteName");

                if (string.IsNullOrEmpty(item.TrackStatus))
                    item.TrackStatus = "unchecked";
            }
            catch (Exception)
            {
            }

            return item;
        }

        private static string GetNodeText(XmlNode parent, string childName)
        {
            XmlNode node = parent.SelectSingleNode(childName);
            if (node != null)
                return node.InnerText ?? "";
            return "";
        }
    }

    // ============================
    // Check Result
    // ============================

    public class CheckResult
    {
        public string ProgramName { get; set; } = "";
        public string Status { get; set; } = "";
        public string Note { get; set; } = "";
        public string Hash { get; set; } = "";
        public string TrackBlock { get; set; } = "";
        public string ExtractedVersion { get; set; } = "";
    }
}
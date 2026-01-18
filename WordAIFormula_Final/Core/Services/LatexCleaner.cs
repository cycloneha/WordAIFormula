using System;
using System.Text.RegularExpressions;

namespace WordAddAIFormula_Final.Services
{
    public class LatexCleaner
    {
        private static readonly Regex DelimiterRegex = new Regex(
            @"^\s*(?<d>\$\$|\\\[|\\\(|\\begin\{equation\}|\\begin\{align\}|\\begin\{eqnarray\})(?<content>.*?)(?<!\\)\k<d>(?=\s*$)",
            RegexOptions.Singleline | RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static string CleanLatex(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            var trimmedInput = input.Trim();

            var match = DelimiterRegex.Match(trimmedInput);

            if (match.Success)
            {
                var content = match.Groups["content"].Value;

                if (trimmedInput.StartsWith("$$") && trimmedInput.EndsWith("$$") && trimmedInput.Length >= 4)
                {
                    content = trimmedInput.Substring(2, trimmedInput.Length - 4);
                }
                else if (trimmedInput.StartsWith("$") && trimmedInput.EndsWith("$") && trimmedInput.Length >= 2)
                {
                    content = trimmedInput.Substring(1, trimmedInput.Length - 2);
                }

                return content.Trim();
            }

            if (trimmedInput.StartsWith("$$") && trimmedInput.EndsWith("$$"))
            {
                return trimmedInput.Substring(2, trimmedInput.Length - 4).Trim();
            }

            if (trimmedInput.StartsWith("$") && trimmedInput.EndsWith("$") && trimmedInput.Length > 2)
            {
                return trimmedInput.Substring(1, trimmedInput.Length - 2).Trim();
            }

            return input;
        }

        public static bool ContainsLatexDelimiters(string input)
        {
            if (string.IsNullOrEmpty(input))
                return false;

            var trimmedInput = input.Trim();

            return
                (trimmedInput.StartsWith("$$") && trimmedInput.EndsWith("$$")) ||
                (trimmedInput.StartsWith("\\[") && trimmedInput.EndsWith("\\]")) ||
                (trimmedInput.StartsWith("\\(") && trimmedInput.EndsWith("\\)")) ||
                (trimmedInput.StartsWith("\\begin{equation}") && trimmedInput.EndsWith("\\end{equation}")) ||
                (trimmedInput.StartsWith("\\begin{align}") && trimmedInput.EndsWith("\\end{align}")) ||
                (trimmedInput.StartsWith("\\begin{eqnarray}") && trimmedInput.EndsWith("\\end{eqnarray}")) ||
                (trimmedInput.StartsWith("$") && trimmedInput.EndsWith("$") && trimmedInput.Length > 2);
        }
    }
}
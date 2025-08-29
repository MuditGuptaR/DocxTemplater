using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;

namespace DocxTemplater
{
    internal class MergeFieldParser
    {
        private static readonly Regex MergeFieldRegex = new Regex(@"^\s*MERGEFIELD\s+""?([^""\s]+)""?\s*\\?\*?\s*[MERGEFORMAT|CHARFORMAT|UPPER|LOWER|FIRSTCAP|CAPS]*\s*$", RegexOptions.IgnoreCase);
        private static readonly Regex IfFieldRegex = new Regex(@"^\s*IF\s+""?([^""\s]+)""?\s+""([^""]*)""\s+""([^""]*)""\s*$", RegexOptions.IgnoreCase);
        private static readonly Regex IncludePictureRegex = new Regex(@"^\s*INCLUDEPICTURE\s+""MERGEFIELD\s+([^""]+)""\s*$", RegexOptions.IgnoreCase);

        public IEnumerable<FieldPattern> Parse(OpenXmlCompositeElement root)
        {
            var fields = new List<FieldPattern>();

            // Find all simple fields
            foreach (var simpleField in root.Descendants<SimpleField>())
            {
                var instruction = simpleField.Instruction.InnerText;
                var match = MergeFieldRegex.Match(instruction);
                if (match.Success)
                {
                    var fieldName = match.Groups[1].Value;
                    fields.Add(new MergeFieldPattern(fieldName, instruction, simpleField));
                    continue;
                }

                match = IfFieldRegex.Match(instruction);
                if (match.Success)
                {
                    var condition = match.Groups[1].Value;
                    var trueText = match.Groups[2].Value;
                    var falseText = match.Groups[3].Value;
                    fields.Add(new IfFieldPattern(condition, trueText, falseText, instruction, simpleField));
                    continue;
                }

                match = IncludePictureRegex.Match(instruction);
                if (match.Success)
                {
                    var pathFieldName = match.Groups[1].Value;
                    fields.Add(new IncludePicturePattern(pathFieldName, instruction, simpleField));
                }
            }

            // Find all complex fields
            var fieldChars = root.Descendants<FieldChar>().ToList();
            var beginChars = fieldChars.Where(f => f.FieldCharType == FieldCharValues.Begin).ToList();

            foreach (var beginChar in beginChars)
            {
                var endChar = FindMatchingEndChar(beginChar);
                if (endChar != null)
                {
                    var fieldCodeElement = FindFieldCode(beginChar, endChar);
                    if (fieldCodeElement != null)
                    {
                        var fieldCode = fieldCodeElement.InnerText;
                        var match = MergeFieldRegex.Match(fieldCode);
                        if (match.Success)
                        {
                            var fieldName = match.Groups[1].Value;
                            fields.Add(new MergeFieldPattern(fieldName, fieldCode, beginChar.Parent));
                            continue;
                        }

                        match = IfFieldRegex.Match(fieldCode);
                        if (match.Success)
                        {
                            var condition = match.Groups[1].Value;
                            var trueText = match.Groups[2].Value;
                            var falseText = match.Groups[3].Value;
                            fields.Add(new IfFieldPattern(condition, trueText, falseText, fieldCode, beginChar.Parent));
                            continue;
                        }

                        match = IncludePictureRegex.Match(fieldCode);
                        if (match.Success)
                        {
                            var pathFieldName = match.Groups[1].Value;
                            fields.Add(new IncludePicturePattern(pathFieldName, fieldCode, beginChar.Parent));
                        }
                    }
                }
            }

            return fields;
        }

        private FieldCode FindFieldCode(FieldChar beginChar, FieldChar endChar)
        {
            var elementsBetween = new List<OpenXmlElement>();
            var current = beginChar.Parent.NextSibling();
            while (current != null && current != endChar.Parent)
            {
                elementsBetween.Add(current);
                current = current.NextSibling();
            }
            return elementsBetween.OfType<Run>().SelectMany(r => r.Elements<FieldCode>()).FirstOrDefault();
        }

        private FieldChar FindMatchingEndChar(FieldChar beginChar)
        {
            // Simplified implementation: does not handle nested fields.
            var sibling = beginChar.Parent.NextSibling();
            while (sibling != null)
            {
                var endChar = sibling.Descendants<FieldChar>().FirstOrDefault(f => f.FieldCharType == FieldCharValues.End);
                if (endChar != null)
                {
                    return endChar;
                }
                var begin = sibling.Descendants<FieldChar>().FirstOrDefault(f => f.FieldCharType == FieldCharValues.Begin);
                if (begin != null)
                {
                    // Found a nested field, skip it.
                    var nestedEnd = FindMatchingEndChar(begin);
                    if (nestedEnd != null)
                    {
                        sibling = nestedEnd.Parent;
                    }
                }
                sibling = sibling.NextSibling();
            }
            return null;
        }
    }
}

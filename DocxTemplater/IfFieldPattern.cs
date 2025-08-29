using DocumentFormat.OpenXml;

namespace DocxTemplater
{
    internal class IfFieldPattern : FieldPattern
    {
        public string Condition { get; }
        public string TrueText { get; }
        public string FalseText { get; }

        public IfFieldPattern(string condition, string trueText, string falseText, string fieldCode, OpenXmlElement element)
            : base(fieldCode, element)
        {
            Condition = condition;
            TrueText = trueText;
            FalseText = falseText;
        }
    }
}

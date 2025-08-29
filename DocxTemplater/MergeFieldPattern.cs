using DocumentFormat.OpenXml;

namespace DocxTemplater
{
    internal class MergeFieldPattern : FieldPattern
    {
        public string FieldName { get; }

        public MergeFieldPattern(string fieldName, string fieldCode, OpenXmlElement element)
            : base(fieldCode, element)
        {
            FieldName = fieldName;
        }
    }
}

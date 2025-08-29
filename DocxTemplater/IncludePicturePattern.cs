using DocumentFormat.OpenXml;

namespace DocxTemplater
{
    internal class IncludePicturePattern : FieldPattern
    {
        public string PathFieldName { get; }

        public IncludePicturePattern(string pathFieldName, string fieldCode, OpenXmlElement element)
            : base(fieldCode, element)
        {
            PathFieldName = pathFieldName;
        }
    }
}

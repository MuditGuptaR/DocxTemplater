using DocumentFormat.OpenXml;

namespace DocxTemplater
{
    internal abstract class FieldPattern
    {
        public string FieldCode { get; }
        public OpenXmlElement Element { get; }

        protected FieldPattern(string fieldCode, OpenXmlElement element)
        {
            FieldCode = fieldCode;
            Element = element;
        }
    }
}

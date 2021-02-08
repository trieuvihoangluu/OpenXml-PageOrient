using System;
using DocumentFormat.OpenXml;

namespace OpenXml
{
    public static class Extensions

    {
        public const  string BioContentId = "bioContentId";
        public static void SetAttribute(this OpenXmlElement element, string x, string y)
        {
            element.SetAttribute(new OpenXmlAttribute(x, "", y));
        }
        public static string GetAttribute(this OpenXmlElement element, string x) => element.GetAttribute(x, "").Value;

        public static int GetBioAttribute(this OpenXmlElement element) =>
            Convert.ToInt32(element.GetAttribute(BioContentId));
    }
}
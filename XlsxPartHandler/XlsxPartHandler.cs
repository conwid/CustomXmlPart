using Ionic.Zip;
using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;

namespace XlsxPartHandler
{
    /// <summary>
    /// Base class to convert xlsx files to resx or resx to xlsx
    /// </summary>
    public static class XlsxPartHandler
    {

        private const string customPartFolder = "/customXml/";
        private const string customXmlName = "item1.xml";


        #region private const string xmlPropertiesContent
        private const string xmlPropertiesContent =
                "<?xml version=\"1.0\" encoding =\"UTF-8\" standalone=\"no\"?>" +
                    "<ds:datastoreItem ds:itemID = \"{" + "{0}" + "}\" xmlns:ds = \"http://schemas.openxmlformats.org/officeDocument/2006/customXml\">" +
                "<ds:schemaRefs/>" +
             "</ds:datastoreItem>";
        #endregion
        private const string workbookPackagePartUri = "/xl/workbook.xml";

        private const string customXmlRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
        private const string customXmlRelationshipPropertiesType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";

        private const string contentTypeXmlPath = "[Content_Types].xml";
        private const string customXmlPropertiesName = "itemProps1.xml";

        private const string customXmlType = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
        private const string typedNodeElementName = "/Types";

        public static void AddCustomXmlPart(string excelFilePath, string xmlFilePath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFilePath);
            var excelFile = new FileInfo(excelFilePath);

            AddCustomParts(xmlDoc, excelFile);
            UpdatePackageContentFile(excelFile);
        }

        private static XmlDocument GetProperties()
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(string.Format(xmlPropertiesContent, Guid.NewGuid().ToString()));
            return xmlDoc;
        }

        private static void AddCustomParts(XmlDocument doc, FileInfo excelFile)
        {
            using (Package package = Package.Open(excelFile.FullName, FileMode.Open, FileAccess.ReadWrite))
            {
                AddCustomParts(doc, package);
            }
        }

        private static void AddCustomParts(XmlDocument doc, Package package)
        {
            package.GetPart(new Uri(workbookPackagePartUri, UriKind.Relative))
                .CreateRelationship(new Uri(Path.Combine("..", customPartFolder, customXmlName), UriKind.Relative), TargetMode.Internal, customXmlRelationshipType);
            AddCustomXmlPropertiesPart(package);
            AddCustomXmlPart(package, doc);
        }

        private static void AddCustomXmlPart(Package package, XmlDocument doc)
        {
            var customXmlContent = CreateCustomXmlPart(new Uri(Path.Combine(customPartFolder, customXmlName), UriKind.Relative), package, doc); // create the part
            customXmlContent.CreateRelationship(new Uri(customXmlPropertiesName, UriKind.Relative), TargetMode.Internal, customXmlRelationshipPropertiesType); //add a relationship in the custom xml part to the properties part
        }
        private static void AddCustomXmlPropertiesPart(Package package)
        {
            XmlDocument xmlDocProperties = GetProperties();
            CreateCustomXmlPart(new Uri(Path.Combine(customPartFolder, customXmlPropertiesName), UriKind.Relative), package, xmlDocProperties);
        }
        private static PackagePart CreateCustomXmlPart(Uri partUri, Package package, XmlDocument partContent)
        {
            PackagePart customXmlContent = package.CreatePart(partUri, customXmlType);
            using (Stream partStream = customXmlContent.GetStream(FileMode.Create, FileAccess.ReadWrite))
            {
                partContent.Save(partStream);
            }
            return customXmlContent;
        }
        private static void UpdatePackageContentFile(FileInfo excelFile)
        {
            using (ZipFile f = new ZipFile(excelFile.FullName))
            {
                var t = f.Entries.Single(e => e.FileName == contentTypeXmlPath);
                XmlDocument d = new XmlDocument();
                using (MemoryStream ms = new MemoryStream())
                {
                    t.Extract(ms);
                    ms.Seek(0, SeekOrigin.Begin);
                    d.Load(ms);
                    var typesNode = d.DocumentElement;
                    var customXmlNode = typesNode.ChildNodes.Cast<XmlNode>().SingleOrDefault(e => e.Attributes["PartName"]?.Value == Path.Combine(customPartFolder, customXmlName)); // this node should be deleted (if it exists) in order for Excel to keep the custom part after saving the document
                    if (customXmlNode != null)
                    {
                        typesNode.RemoveChild(customXmlNode);
                    }
                }
                f.UpdateEntry(contentTypeXmlPath, d.OuterXml);
                f.Save();
            }
        }
    }
}
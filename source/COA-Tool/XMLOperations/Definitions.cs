using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;

namespace CoA_Tool.XMLOperations
{
    /// <summary>
    /// Contains operations relating to XML files within the definitions directory.
    /// </summary>
    static class Definitions
    {
        public enum DefinitionFiles { Micro};
        /// <summary>
        /// Retrieves a node's inner text via an XPath expression.
        /// </summary>
        /// <param name="defFile">The file's name as an enum.</param>
        /// <param name="xPathExpression">The XPath expression for the target node.</param>
        /// <returns></returns>
        public static string NodeValueViaXPath(DefinitionFiles defFile , string xPathExpression)
        {
            QueuedAccess queuedAccess = new QueuedAccess(Directory.GetCurrentDirectory() + "/Definitions/" + Convert.ToString(defFile) + ".xml", 
                FileMode.Open, FileAccess.Read, FileShare.Read);
            using (queuedAccess)
            {
                XmlDocument document = new XmlDocument();
                document.Load(queuedAccess.XMLFileStream);
                return document.SelectSingleNode(xPathExpression).InnerText.Replace('"', ' ');
            }
        }


    }
}

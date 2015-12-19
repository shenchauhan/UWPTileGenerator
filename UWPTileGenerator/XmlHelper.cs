using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace UWPTileGenerator
{
	public static class XmlHelper
	{
		public static void AddAttribute(this XElement element, string attribute, string value)
		{
			var xAttribute = element.Attribute(XName.Get(attribute));

			if (xAttribute == null)
			{
				element.Add(new XAttribute(XName.Get(attribute), value));
			}
			else
			{
				element.Attribute(XName.Get(attribute)).Value = value;
			}
		}
	}
}

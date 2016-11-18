/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.IO;
    using System.Text;
    using System.Xml;
    using System.Xml.Serialization;

    /// <remarks/>
    [Serializable]
    [XmlType(AnonymousType = true, Namespace = "CategoryList.xsd")]
    [XmlRoot(ElementName = "categories", Namespace = "CategoryList.xsd", IsNullable = false)]
    public class CategoryList
    {
        private UserConfiguration _UserConfigurationItem;

        /// <remarks/>
        [XmlElement("category")]
        public List<Category> Categories { get; set; }

        /// <remarks/>
        [XmlIgnore]
        public Guid? DefaultCategory { get; set; }

        [XmlAttribute("default")]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public string DefaultCategoryText
        {
            get { return DefaultCategory != null ? DefaultCategory.ToString() : string.Empty; }
            set
            {
#if net40
                Guid result;
                DefaultCategory = (Guid.TryParse(value, out result)) ? result : (Guid?)null;
#else
                try
                {
                    DefaultCategory = new Guid(value);
                }
                catch (Exception)
                {
                    DefaultCategory = null;
                }
#endif
            }
        }


        /// <remarks/>
        [XmlAttribute("lastSavedSession")]
        public byte LastSavedSession { get; set; }

        /// <remarks/>
        [XmlAttribute("lastSavedTime")]
        public DateTime LastSavedTime { get; set; }

        public static CategoryList Bind(ExchangeService service)
        {
            var item = UserConfiguration.Bind(service, 
                "CategoryList", 
                WellKnownFolderName.Calendar,                               
                UserConfigurationProperties.XmlData);

            using (var memory = new MemoryStream(item.XmlData))
            {
                var reader = new StreamReader(memory, Encoding.UTF8, true);
                var serializer = new XmlSerializer(typeof(CategoryList));
                var result = (CategoryList)serializer.Deserialize(reader);
                result._UserConfigurationItem = item;
                return result;
            }
        }

        public void Update()
        {
            using (var stream = new MemoryStream())
            {
                var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = Encoding.UTF8 });
                var serializer = new XmlSerializer(typeof(CategoryList));

                serializer.Serialize(writer, this);
                writer.Flush();
                _UserConfigurationItem.XmlData = stream.ToArray();
                _UserConfigurationItem.Update();
            }
        }
    }
}

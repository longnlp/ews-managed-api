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
    using System.Globalization;
    using System.IO;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// XML reader.
    /// </summary>
    internal class EwsServiceXmlReader : EwsXmlReader
    {
        #region Private members

        private ExchangeService service;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="EwsServiceXmlReader"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="service">The service.</param>
        internal EwsServiceXmlReader(Stream stream, ExchangeService service)
            : base(stream)
        {
            this.service = service;
        }

        #endregion

        /// <summary>
        /// Converts the specified string into a DateTime objects.
        /// </summary>
        /// <param name="dateTimeString">The date time string to convert.</param>
        /// <returns>A DateTime representing the converted string.</returns>
        private DateTime? ConvertStringToDateTime(string dateTimeString)
        {
            return this.Service.ConvertUniversalDateTimeStringToLocalDateTime(dateTimeString);
        }

        /// <summary>
        /// Converts the specified string into a unspecified Date object, ignoring offset.
        /// </summary>
        /// <param name="dateTimeString">The date time string to convert.</param>
        /// <returns>A DateTime representing the converted string.</returns>
        private DateTime? ConvertStringToUnspecifiedDate(string dateTimeString)
        {
            return this.Service.ConvertStartDateToUnspecifiedDateTime(dateTimeString);
        }

        /// <summary>
        /// Reads the element value as date time.
        /// </summary>
        /// <returns>Element value.</returns>
        public DateTime? ReadElementValueAsDateTime()
        {
            return this.ConvertStringToDateTime(this.ReadElementValue());
        }

        /// <summary>
        /// Reads the element value as unspecified date.
        /// </summary>
        /// <returns>Element value.</returns>
        public DateTime? ReadElementValueAsUnspecifiedDate()
        {
            return this.ConvertStringToUnspecifiedDate(this.ReadElementValue());
        }

        /// <summary>
        /// Reads the element value as date time, assuming it is unbiased (e.g. 2009/01/01T08:00) 
        /// and scoped to service's time zone.
        /// </summary>
        /// <returns>The element's value as a DateTime object.</returns>
        public DateTime ReadElementValueAsUnbiasedDateTimeScopedToServiceTimeZone()
        {
            string elementValue = this.ReadElementValue();
            return EwsUtilities.ParseAsUnbiasedDatetimescopedToServicetimeZone(elementValue, this.Service);
        }

        /// <summary>
        /// Reads the element value as date time.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">Name of the local.</param>
        /// <returns>Element value.</returns>
        public DateTime? ReadElementValueAsDateTime(XmlNamespace xmlNamespace, string localName)
        {
            return this.ConvertStringToDateTime(this.ReadElementValue(xmlNamespace, localName));
        }

        /// <summary>
        /// Reads the service objects collection from XML.
        /// </summary>
        /// <typeparam name="TServiceObject">The type of the service object.</typeparam>
        /// <param name="collectionXmlNamespace">Namespace of the collection XML element.</param>
        /// <param name="collectionXmlElementName">Name of the collection XML element.</param>
        /// <param name="getObjectInstanceDelegate">The get object instance delegate.</param>
        /// <param name="clearPropertyBag">if set to <c>true</c> [clear property bag].</param>
        /// <param name="requestedPropertySet">The requested property set.</param>
        /// <param name="summaryPropertiesOnly">if set to <c>true</c> [summary properties only].</param>
        /// <returns>List of service objects.</returns>
        public List<TServiceObject> ReadServiceObjectsCollectionFromXml<TServiceObject>(
            XmlNamespace collectionXmlNamespace,
            string collectionXmlElementName,
            GetObjectInstanceDelegate<TServiceObject> getObjectInstanceDelegate,
            bool clearPropertyBag,
            PropertySet requestedPropertySet,
            bool summaryPropertiesOnly) where TServiceObject : ServiceObject
        {
            List<TServiceObject> serviceObjects = new List<TServiceObject>();
            TServiceObject serviceObject = null;

            if (!this.IsStartElement(collectionXmlNamespace, collectionXmlElementName))
            {
                this.ReadStartElement(collectionXmlNamespace, collectionXmlElementName);
            }

            if (!this.IsEmptyElement)
            {
                do
                {
                    this.Read();

                    if (this.IsStartElement())
                    {
                        serviceObject = getObjectInstanceDelegate(this.Service, this.LocalName);

                        if (serviceObject == null)
                        {
                            this.SkipCurrentElement();
                        }
                        else
                        {
                            if (!ProcessTypeMatch(LocalName, serviceObject))
                            {
                                throw new ServiceLocalException(
                                    string.Format(
                                        "The type of the object in the store ({0}) does not match that of the local object ({1}).",
                                        this.LocalName,
                                        serviceObject.GetXmlElementName()));
                            }

                            serviceObject.LoadFromXml(
                                            this,
                                            clearPropertyBag,
                                            requestedPropertySet,
                                            summaryPropertiesOnly);

                            serviceObjects.Add(serviceObject);
                        }
                    }
                }
                while (!this.IsEndElement(collectionXmlNamespace, collectionXmlElementName));
            }

            return serviceObjects;
        }

        /// <summary>
        /// Reads the service objects collection from XML.
        /// </summary>
        /// <typeparam name="TServiceObject">The type of the service object.</typeparam>
        /// <param name="collectionXmlElementName">Name of the collection XML element.</param>
        /// <param name="getObjectInstanceDelegate">The get object instance delegate.</param>
        /// <param name="clearPropertyBag">if set to <c>true</c> [clear property bag].</param>
        /// <param name="requestedPropertySet">The requested property set.</param>
        /// <param name="summaryPropertiesOnly">if set to <c>true</c> [summary properties only].</param>
        /// <returns>List of service objects.</returns>
        public List<TServiceObject> ReadServiceObjectsCollectionFromXml<TServiceObject>(
            string collectionXmlElementName,
            GetObjectInstanceDelegate<TServiceObject> getObjectInstanceDelegate,
            bool clearPropertyBag,
            PropertySet requestedPropertySet,
            bool summaryPropertiesOnly) where TServiceObject : ServiceObject
        {
            return this.ReadServiceObjectsCollectionFromXml<TServiceObject>(
                                XmlNamespace.Messages,
                                collectionXmlElementName,
                                getObjectInstanceDelegate,
                                clearPropertyBag,
                                requestedPropertySet,
                                summaryPropertiesOnly);
        }

        /// <summary>
        /// Gets the service.
        /// </summary>
        /// <value>The service.</value>
        public ExchangeService Service
        {
            get { return this.service; }
        }

        private bool ProcessTypeMatch<TServiceObject>(string typeName, TServiceObject serviceObject)
            where TServiceObject : ServiceObject
        {
            var matched = false;

            if (string.Compare(this.LocalName, serviceObject.GetXmlElementName(), StringComparison.Ordinal) != 0)
            {
                if (service.ReadCompatibleServiceObject)
                {
                    matched = IsTypeCompatibleWithServiceObject(typeName, serviceObject);
                }
            }
            else
            {
                matched = true;
            }

            return matched;
        }

        /// <summary>
        /// Get the compatible from the type name with the service object type
        /// </summary>
        /// <param name="typeName"></param>
        /// <param name="serviceObjectType"></param>
        /// <returns></returns>
        private static bool IsTypeCompatibleWithServiceObject<TServiceObject>(string typeName, TServiceObject serviceObject)
            where TServiceObject : ServiceObject
        {
            var index = new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase)
            {
                { "MeetingMessage", typeof(MeetingMessage) },
                { "MeetingCancellation", typeof(MeetingCancellation) },
                { "MeetingRequest", typeof(MeetingRequest) },
                { "MeetingResponse", typeof(MeetingResponse) },
            };

            Type type;

            if (index.TryGetValue(typeName, out type))
            {
                if (serviceObject.GetType().IsAssignableFrom(type))
                {
                    serviceObject.SetXmlElementName(typeName);

                    return true;
                }
            }

            return false;
        }
    }
}
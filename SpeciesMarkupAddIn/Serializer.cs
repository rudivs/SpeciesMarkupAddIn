using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.Serialization;
using System.Xml.Serialization;
using System.Xml;

namespace SpeciesMarkupAddIn
{

    public class Serializer
    {
        public Serializer()
        {

        }

        public void SerializeObject(string filename, TaxonList objectToSerialize)
        {
            XmlWriterSettings ws = new XmlWriterSettings();
            ws.NewLineHandling = NewLineHandling.Entitize;

            XmlSerializer serializer = new XmlSerializer(typeof(TaxonList));
            using (XmlWriter wr = XmlWriter.Create(filename, ws))
            {
                serializer.Serialize(wr, objectToSerialize);
            }
        }

        public TaxonList DeSerializeObject(string filename)
        {
            TaxonList objectToSerialize;
            Stream stream = File.Open(filename, FileMode.Open);
            XmlSerializer serializer = new XmlSerializer(typeof(TaxonList));
            objectToSerialize = (TaxonList)serializer.Deserialize(stream);
            stream.Close();
            return objectToSerialize;
        }
    }
}
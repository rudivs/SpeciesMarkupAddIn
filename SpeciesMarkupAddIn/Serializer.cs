using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace SpeciesMarkupAddIn
{

    public class Serializer
    {
        public Serializer()
        {

        }

        public void SerializeObject(string filename, TaxonList objectToSerialize)
        {
            Stream stream = File.Open(filename, FileMode.Create);
            XmlSerializer serializer = new XmlSerializer(typeof(TaxonList));
            serializer.Serialize(stream, objectToSerialize);
            stream.Close();
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
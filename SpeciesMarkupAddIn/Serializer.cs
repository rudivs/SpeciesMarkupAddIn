using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

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
            BinaryFormatter bFormatter = new BinaryFormatter();
            bFormatter.Serialize(stream, objectToSerialize);
            stream.Close();
        }

        public TaxonList DeSerializeObject(string filename)
        {
            TaxonList objectToSerialize;
            Stream stream = File.Open(filename, FileMode.Open);
            BinaryFormatter bFormatter = new BinaryFormatter();
            objectToSerialize = (TaxonList)bFormatter.Deserialize(stream);
            stream.Close();
            return objectToSerialize;
        }
    }
}
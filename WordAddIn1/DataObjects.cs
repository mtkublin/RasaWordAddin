using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        // TRAIN -------------------------------------------------------------------------------------------------

        private class Ent
        {
            public int start { get; set; }
            public int end { get; set; }
            public string value { get; set; }
            public string entity { get; set; }

            public Ent(int startchar, int endchar, string val, string ent)
            {
                start = startchar;
                end = endchar;
                value = val;
                entity = ent;
            }
        }

        private class Examp
        {
            public string text { get; set; }
            public string intent { get; set; }
            public List<Ent> entities { get; set; }

            public Examp(string sentText, string intnt, List<Ent> entities1)
            {
                text = sentText;
                intent = intnt;
                entities = entities1;
            }
        }

        private class TrainData
        {
            public List<Examp> common_examples { get; set; }
            public List<object> regex_features { get; set; }
            public List<object> lookup_tables { get; set; }
            public List<object> entity_synonyms { get; set; }

            public TrainData(List<Examp> examps)
            {
                common_examples = examps;
                regex_features = new List<object>();
                lookup_tables = new List<object>();
                entity_synonyms = new List<object>();
            }
        }

        private class RasaNLUdata
        {
            public TrainData rasa_nlu_data { get; set; }
            public string ModelPath { get; set; }

            public RasaNLUdata(TrainData DataToPass, string Mpath = null)
            {
                rasa_nlu_data = DataToPass;
                if (Mpath != null)
                {
                    ModelPath = Mpath;
                }
            }
        }

        private class FinalDataObject
        {
            public RasaNLUdata DATA { get; set; }

            public FinalDataObject(RasaNLUdata DataToPass)
            {
                DATA = DataToPass;
            }
        }

        // TEST -------------------------------------------------------------------------------------------------

        private class TextToExportObject
        {
            public List<String> SENTS { get; set; }

            public TextToExportObject(List<String> DataToPass)
            {
                SENTS = DataToPass;
            }
        }

        private class FinalTestDataExportObject
        {
            public TextToExportObject DATA { get; set; }

            public FinalTestDataExportObject(TextToExportObject DataToPass)
            {
                DATA = DataToPass;
            }
        }

        public class Intent
        {
            public string name { get; set; }
            public double confidence { get; set; }
        }

        public class SingleEnt
        {
            public int start { get; set; }
            public int end { get; set; }
            public string value { get; set; }
            public string entity { get; set; }
            public float confidence { get; set; }
            public string extractor { get; set; }
        }

        public class IntentRanking
        {
            public string name { get; set; }
            public double confidence { get; set; }
        }

        public class SentenceObject
        {
            public Intent intent { get; set; }
            public List<SingleEnt> entities { get; set; }
            public List<IntentRanking> intent_ranking { get; set; }
            public string text { get; set; }
        }

        public class TestDataReqIDobject
        {
            public string req_id { get; set; }
            public string mongo_id { get; set; }
            public string status { get; set; }

        }

        // OTHER -------------------------------------------------------------------------------------------------

        private class ModelPathDataObject
        { 
            public string DATA { get; set; }

            public ModelPathDataObject(string model_pathData)
            {
                DATA = model_pathData;
            }
        }
    }
}

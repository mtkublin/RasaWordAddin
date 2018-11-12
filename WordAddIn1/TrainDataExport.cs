﻿using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        public void ExportTrainData()
        {
            var examps = new List<Examp> { };

            foreach (ContentControl intent in Globals.ThisAddIn.Application.ActiveDocument.ContentControls)
            {
                string intTag = intent.Tag;
                char intLevelIndicator = intTag[intTag.Length - 1];

                if (intLevelIndicator is '1')
                {
                    GatherTrainData(intTag, intent.Range, examps);

                    TrainData tData = new TrainData(examps);
                    string outputJSON = "{\"rasa_nlu_data\":" + JsonConvert.SerializeObject(tData) + "}";

                    string mydocpath = @"C:\Users\Mikołaj\Documents\Word_rasa_addin";
                    using (StreamWriter outputFile = new StreamWriter(Path.Combine(mydocpath, "TrainData.json")))
                    {
                        outputFile.WriteLine(outputJSON);
                    }
                }
            }
        }

        private void GatherTrainData(string intTag, Range sent, List<Examp> examps)
        {
            string sentText = sent.Text;
            string sentInt = intTag;
            int intentStart = sent.Start;

            var entities = new List<Ent> { };
            int EntNumber = 0;
            foreach (ContentControl ent in sent.ContentControls)
            {
                string entTag = ent.Tag;
                char entLevelIndicator = entTag[entTag.Length - 1];

                if (entLevelIndicator is '2')
                {
                    int st = ent.Range.Start - intentStart - 1 - EntNumber;
                    int en = ent.Range.End - intentStart - 1 - EntNumber;
                    string val = ent.Range.Text;
                    string tag = entTag;

                    Ent entity = new Ent(st, en, val, tag);
                    entities.Add(entity);

                    EntNumber += 2;
                }
            }
            Examp examp = new Examp(sentText, sentInt, entities);
            examps.Add(examp);
        }

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

            public TrainData(List<Examp> examps)
            {
                common_examples = examps;
            }
        }
    }
}
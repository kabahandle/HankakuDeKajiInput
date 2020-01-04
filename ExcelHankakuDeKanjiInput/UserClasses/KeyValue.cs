using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HankakuDeKanjiInput
{
    public class KeyValue
    {
        public string Key = string.Empty;
        public string Value = string.Empty;
        public KeyValue(string key, string value)
        {
            this.Key = key;
            this.Value = value;
        }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HankakuDeKanjiInput
{
    public class WordItem
    {
        public string segment = string.Empty;
        public string word = string.Empty;
        public WordItem(string sg, string txt)
        {
            this.segment = sg;
            this.word = txt;
        }
        public override string ToString()
        {
            return this.word;
        }

    }
}

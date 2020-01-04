using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualBasic;
using System.Web;
using System.IO;
using System.Net;
using System.Windows.Forms;
using HankakuDeKanjiInput;

namespace MyVSNetAddin2
{
    public class YahooAPI
    {
        static string[] kn = new string[100];
        static string[] al = new string[100];

        public static WebBrowser webTrans = null;
        private string[] yahooAppID = null;//new string[10];
        private static CtrlWebTrans ctrlWebTrans = new CtrlWebTrans();
        public YahooAPI()
        {
            YahooAPI.webTrans = YahooAPI.ctrlWebTrans.webTrans;

            FileStream fs = null;
            StreamReader r = null;
            List<string> userAPIIDs = new List<string>();
            try
            {
                fs = File.Open(Application.StartupPath + @"\App_Data\yahooApiIDs.txt", FileMode.OpenOrCreate);
                r = new StreamReader(fs, Encoding.Default);

                string line = "";
                while ((line = r.ReadLine()) != null)
                {
                    if (!string.IsNullOrEmpty(line.Trim()))
                    {
                        userAPIIDs.Add(line);
                    }
                }

                fs.Close();
            }
            catch (Exception exc)
            {
            }
            finally
            {
                if (r != null)
                {
                    r.Close();
                }
                if (fs != null)
                {
                    fs.Close();
                }
            }

            if (userAPIIDs.Count < 1)
            {
                yahooAppID = new string[10];
                yahooAppID[0] = "._AeiPuxg67zicSBZR4epBvPSC3_dYB6DRE4e8_2PWZY.lmlky6YzlUDJiPVu9JdnA--";
                yahooAppID[1] = "0BmlIBSxg65rCL5dt_MKkOqKQJmajLswknxmjtMBd_JnLvKFcw0nElk_VxwZ.wqFwA--";
                yahooAppID[2] = "MowY2WCxg64xjYYyFK4KeYAkqSfAIkPmUnQQgw44KXovWTlUWRUKOJsFJuyVDlY91A--";
                yahooAppID[3] = "H80H8H.xg66K8LOnhk3ap_.2y36XS2sTv1wo0ZPmSiQ1DcCy0fMZdFGl1TQ8AJmyWw--";
                yahooAppID[4] = "wrjUJOmxg65P.L_.Umxq3vShwM2fWp838arZC5vvTuerpQnSfHx3FRpeRvpi3JwmYg--";
                yahooAppID[5] = "SV9sHZ2xg67GdDcoSLNgUER0Qymwdr7H25ynvBv_nqlwilQUk74I7b2lugkV72F3Iw--";
                yahooAppID[6] = "Ww6iVq.xg67x6_b82MZFxiRdubh3xN71X6z4F5QIMoNYHs288bSKyUM.T9V6nZRs.Q--";
                yahooAppID[7] = "BR9U5hqxg66f84.CL1fembsrV.WJUrvGNgfF3qWiFJXQ3wNDWCPpqFJDnvZ3i.WZlQ--";
                yahooAppID[8] = "r3ylg5uxg64Crt9xAK5t1GB4zBl_ehrivVU6ooozMtzW4o.kPo0pwbMxHSioigtdOg--";
                yahooAppID[9] = "VzWfy5Cxg67VCz31WTkgOfBVFIjNQ6ChcRqgSAiKH40Vk1879tEzFcAEeVaV7Cd2Sw--";
            }
            else
            {
                yahooAppID = userAPIIDs.ToArray();
            }

        }
        public string GetApiID()
        {
            Random rnd = new Random();
            int randomNumber = rnd.Next(yahooAppID.Length);

            return yahooAppID[randomNumber];
        }

        public string GetJPWordForyahooAPI(string jpWord)
        {
            jpWord = Strings.StrConv(jpWord, VbStrConv.Hiragana, 0);

//            string[] kn = new string[100];
//            string[] al = new string[100];

            kn[0] = "あ";
            kn[1] = "い";
            kn[2] = "う";
            kn[3] = "え";
            kn[4] = "お";
            kn[5] = "か";
            kn[6] = "き";
            kn[7] = "く";
            kn[8] = "け";
            kn[9] = "こ";
            kn[10] = "さ";
            kn[11] = "し";
            kn[12] = "す";
            kn[13] = "せ";
            kn[14] = "そ";
            kn[15] = "た";
            kn[16] = "ち";
            kn[17] = "つ";
            kn[18] = "て";
            kn[19] = "と";
            kn[20] = "な";
            kn[21] = "に";
            kn[22] = "ぬ";
            kn[23] = "ね";
            kn[24] = "の";
            kn[25] = "は";
            kn[26] = "ひ";
            kn[27] = "ふ";
            kn[28] = "へ";
            kn[29] = "ほ";
            kn[30] = "ま";
            kn[31] = "み";
            kn[32] = "む";
            kn[33] = "め";
            kn[34] = "も";
            kn[35] = "や";
            kn[36] = "ゆ";
            kn[37] = "よ";
            kn[38] = "ら";
            kn[39] = "り";
            kn[40] = "る";
            kn[41] = "れ";
            kn[42] = "ろ";
            kn[43] = "わ";
            kn[44] = "を";
            kn[45] = "ん";
            kn[46] = "ぁ";
            kn[47] = "ぃ";
            kn[48] = "ぅ";
            kn[49] = "ぇ";
            kn[50] = "ぉ";
            kn[51] = "ゃ";
            kn[52] = "ゅ";
            kn[53] = "ょ";
            kn[54] = "っ";

            kn[55] = "が";
            kn[56] = "ぎ";
            kn[57] = "ぐ";
            kn[58] = "げ";
            kn[59] = "ご";
            kn[60] = "ざ";
            kn[61] = "じ";
            kn[62] = "ず";
            kn[63] = "ぜ";
            kn[64] = "ぞ";
            kn[65] = "だ";
            kn[66] = "ぢ";
            kn[67] = "づ";
            kn[68] = "で";
            kn[69] = "ど";
            kn[70] = "ば";
            kn[71] = "び";
            kn[72] = "ぶ";
            kn[73] = "べ";
            kn[74] = "ぼ";
            kn[75] = "ぱ";
            kn[76] = "ぴ";
            kn[77] = "ぷ";
            kn[78] = "ぺ";
            kn[79] = "ぽ";


            al[0] = "a";
            al[1] = "i";
            al[2] = "u";
            al[3] = "e";
            al[4] = "o";
            al[5] = "ka";
            al[6] = "ki";
            al[7] = "ku";
            al[8] = "ke";
            al[9] = "ko";
            al[10] = "sa";
            al[11] = "si";
            al[12] = "su";
            al[13] = "se";
            al[14] = "so";
            al[15] = "ta";
            al[16] = "ti";
            al[17] = "tu";
            al[18] = "te";
            al[19] = "to";
            al[20] = "na";
            al[21] = "ni";
            al[22] = "nu";
            al[23] = "ne";
            al[24] = "no";
            al[25] = "ha";
            al[26] = "hi";
            al[27] = "hu";
            al[28] = "he";
            al[29] = "ho";
            al[30] = "ma";
            al[31] = "mi";
            al[32] = "mu";
            al[33] = "me";
            al[34] = "mo";
            al[35] = "ya";
            al[36] = "yu";
            al[37] = "yo";
            al[38] = "ra";
            al[39] = "ri";
            al[40] = "ru";
            al[41] = "re";
            al[42] = "ro";
            al[43] = "wa";
            al[44] = "wo";
            al[45] = "nn";
            al[46] = "la";
            al[47] = "li";
            al[48] = "lu";
            al[49] = "le";
            al[50] = "lo";
            al[51] = "lya";
            al[52] = "lyu";
            al[53] = "lyo";
            al[54] = "ltu";

            al[55] = "ga";
            al[56] = "gi";
            al[57] = "gu";
            al[58] = "ge";
            al[59] = "go";
            al[60] = "za";
            al[61] = "zi";
            al[62] = "zu";
            al[63] = "ze";
            al[64] = "zo";
            al[65] = "da";
            al[66] = "di";
            al[67] = "du";
            al[68] = "de";
            al[69] = "do";
            al[70] = "ba";
            al[71] = "bi";
            al[72] = "bu";
            al[73] = "be";
            al[74] = "bo";
            al[75] = "pa";
            al[76] = "pi";
            al[77] = "pu";
            al[78] = "pe";
            al[79] = "po";

            /*
            for (int i = 0; i < 80; i++)
            {
                jpWord = jpWord.Replace(kn[i], al[i]);
            }
            */
            PrevHiraganaWord = jpWord;
            for (int i = al.Length -1 ; i >= 0; i--)
            {
                if (!string.IsNullOrEmpty(al[i]))
                {
                    PrevHiraganaWord = PrevHiraganaWord.Replace(al[i], kn[i]);
                }
            }

            return jpWord;
        }
        static string PrevHiraganaWord = string.Empty;

        public string getUrlRoman(string transText, string apiid)
        {
            string strgURL = "http://jlp.yahooapis.jp/JIMService/V1/conversion?appid="
                + apiid
                + "&sentence="
                + HttpUtility.UrlEncode(transText)
                + "&format=roman"
                + "&dictionary=default,name,place,zip,symbol";

            return strgURL;
        }

        public string getUrlPredict(string transText, string apiid)
        {
            string strgURL = "http://jlp.yahooapis.jp/JIMService/V1/conversion?appid="
                + apiid
                + "&sentence="
                + HttpUtility.UrlEncode(transText)
                + "&format=roman&mode=predictive";

            return strgURL;
        }

        public List<WordItem> getWords(string apiurl, out List<KeyValue> segments)
        {
            List<WordItem> words = new List<WordItem>();
            segments = new List<KeyValue>();
            /*
            if (!string.IsNullOrEmpty(PrevHiraganaWord))
            {
                words.Add(, PrevHiraganaWord);
            }
            */

            if (YahooAPI.webTrans == null) return words;

            try
            {
                WebClient wc2 = new WebClient();

                Stream st = wc2.OpenRead(apiurl);

                /*
                YahooAPI.webTrans.Navigate(apiurl);
                while (YahooAPI.webTrans.ReadyState != WebBrowserReadyState.Complete)
                {
                    System.Windows.Forms.Application.DoEvents();
                }
                */
                //Stream st = YahooAPI.webTrans.DocumentStream;

                //Encoding enc = Encoding.GetEncoding("Shift_JIS");
                Encoding enc = Encoding.UTF8;
                StreamReader sr = new StreamReader(st, enc);
                string innerText = sr.ReadToEnd();
                sr.Close();

                st.Close();

                string[] lines = innerText.Split('\n');
                /*
                string innerText = YahooAPI.webTrans.Document.Body.InnerText;
                string[] lines = innerText.Split('\n');
                */

                string CurrentSegmentText = string.Empty;
                foreach (string line in lines)
                {
                    string l = line.Trim();
                    if (!string.IsNullOrEmpty(l) && l.StartsWith("<SegmentText>") )
                    {
                        string text = l.Replace("<SegmentText>","").Replace("</SegmentText>","").Trim();
                        if( text != CurrentSegmentText )
                        {
                            CurrentSegmentText = text;
                            if (!string.IsNullOrEmpty(text))
                            {
                                segments.Add(new KeyValue(text, text));
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(l) && l.StartsWith("<Candidate>"))
                    {
                        string word = l.Replace("<Candidate>", "").Replace("</Candidate>", "").Trim();
                        words.Add( new WordItem(CurrentSegmentText, word));
                    }
                }
            }
            catch (Exception exp)
            {
                //throw;
            }
            return words;

        }
    }
}

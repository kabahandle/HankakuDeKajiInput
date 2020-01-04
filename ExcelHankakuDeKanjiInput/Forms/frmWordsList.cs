using Microsoft.VisualBasic;
using MyVSNetAddin2;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace HankakuDeKanjiInput
{
    public partial class frmWordsList : Form
    {
        private string words = string.Empty;
        public string Words
        {
            get
            {
                //return this.words;
                return this.rchText.Text;
            }
            set
            {
                //this.words = value;
                this.rchText.Text = value;
            }
        }
        private string SegmentText { get; set; }
        private List<KeyValue> Segments { get; set; }

        /*
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
         */


        public string StartupPath { get; set; }

        public delegate void WordSelect(string word);
        public event WordSelect OnWordSelect;

        public event Action<Keys> ExitedKeyDown;
        public event Action<char> ExitedKeyPress;

        public string SelectedWord { get; set; }

        public frmWordsList()
        {
            InitializeComponent();
        }

        private void frmWordsList_Load(object sender, EventArgs e)
        {
            this.DataLoad();

        }

        private void MoveFocusToDefault()
        {
            this.tabControl1.Focus();
            if (firstTabPage != null)
            {
                firstTabPage.Focus();
            }
            if (firstListBox != null)
            {
                firstListBox.Focus();
                if (firstListBox.Items.Count > 0)
                {
                    firstListBox.SelectedIndex = 0;
                }
            }
        }


        /*
        public void Activate()
        {
            if (this.Words != null)
            {
                this.lstWords.Items.Clear();
                foreach (string word in this.Words)
                {
                    this.lstWords.Items.Add(word);
                }
            }
        }
         */
        /*
        public void Activate(List<string> words)
        {
            this.Words = words;
            this.Activate();
        }
        */

        private void lstWords_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '\x0d')
            {
                //e.Handled = true;
                //e.KeyChar = '\0';
                //this.Hide();
                return;
            }
            if (this.ExitedKeyPress != null)
            {
                this.ExitedKeyPress(e.KeyChar);
            }
            this.Hide();
        }

        private string MakeReturningString(object sender)
        {
            ListBox list = sender as ListBox;
            if (list == null) return string.Empty;
            WordItem item = list.SelectedItem as WordItem;
            if (item == null) return string.Empty;

            StringBuilder sb = new StringBuilder();
            foreach (KeyValue pair in this.Segments)
            {
                if (!string.IsNullOrEmpty(pair.Key)
                    && !string.IsNullOrEmpty(pair.Value))
                {
                    if (pair.Key == item.segment)
                    {
                        sb.Append(item.word);
                        pair.Value = item.word;
                    }
                    else
                    {
                        sb.Append(pair.Value);
                    }
                }
            }

            /*
            string line = this.SegmentText;
            line = line.Replace(item.segment, item.word);
             */


            //rchText.Text = sb.ToString();
            return sb.ToString();
        }

        private void lstWords_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Control)
            {
                if (this.OnWordSelect != null)
                {
                    this.OnWordSelect(this.rchText.Text);
                }
                this.Hide();
            }
            if (this.ExitedKeyDown != null)
            {
                this.ExitedKeyDown(e.KeyCode);
            }
        }
        private void lstWords_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control)
            {
                /*
                if (this.OnWordSelect != null)
                {
                    this.OnWordSelect(this.rchText.Text);
                }
                this.Hide();
                */
                return;
            }
            if (e.KeyCode == Keys.Enter)
            {
                this.rchText.Text = this.MakeReturningString(sender);
                /*
                ListBox list = sender as ListBox;
                if( list == null ) return;
                WordItem item = list.SelectedItem as WordItem;
                if (item == null) return;

                StringBuilder sb = new StringBuilder();
                foreach (KeyValue pair in this.Segments)
                {
                    if (!string.IsNullOrEmpty(pair.Key)
                        && !string.IsNullOrEmpty(pair.Value))
                    {
                        if (pair.Key == item.segment)
                        {
                            sb.Append(item.word);
                            pair.Value = item.word;
                        }
                        else
                        {
                            sb.Append(pair.Value);
                        }
                    }
                }
                */

                /*
                string line = this.SegmentText;
                line = line.Replace(item.segment, item.word);
                 */


                //rchText.Text = sb.ToString();
                
                /*
                string retWord = sb.ToString();
                if (this.OnWordSelect != null)
                {
                    this.OnWordSelect(retWord);
                }
                 */

                /*
                if (this.lstWords.SelectedItem != null)
                {
                    this.SelectedWord = this.lstWords.SelectedItem.ToString();
                }
                else
                {
                    this.SelectedWord = string.Empty;
                }

                if (this.OnWordSelect != null)
                {
                    this.OnWordSelect(this.SelectedWord);
                }
                //this.DataSave();
                 */
                //this.Hide();
                e.Handled = true;
                return;
            }
            else if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down 
                || e.KeyCode == Keys.Control || e.KeyCode == Keys.Shift || e.KeyCode == Keys.Alt)
            {
                // 通す
                base.OnKeyDown(e);
                //base.OnKeyUp(e);
                return;
            }
            else if (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right)
            {
                ListBox list = sender as ListBox;
                if (list == null) return;
                WordItem item = list.SelectedItem as WordItem;
                if (item == null) return;

                if( e.KeyCode == Keys.Left )
                {
                    if (this.tabControl1.SelectedIndex <= 0)
                    {
                        this.tabControl1.SelectedIndex = this.tabControl1.TabPages.Count - 1;
                    }
                    else
                    {
                        this.tabControl1.SelectedIndex--;
                    }
                }
                else if (e.KeyCode == Keys.Right)
                {
                    if (this.tabControl1.SelectedIndex + 1 >= this.tabControl1.TabPages.Count)
                    {
                        this.tabControl1.SelectedIndex = 0;
                    }
                    else
                    {
                        this.tabControl1.SelectedIndex++;
                    }
                }
                /*
                string segment = item.segment;
                if (string.IsNullOrEmpty(segment)) return;


                int segmax = this.Segments.Count;
                int segcnt = 0;
                int max = this.tabControl1.TabPages.Count;
                int cnt = 0;
                string nextseg = string.Empty;

                if (e.KeyCode == Keys.Right)
                {
                    while (segcnt < segmax)
                    {
                        if (segment.Equals(this.Segments[segcnt]))
                        {
                            int j = segcnt + 1;
                            if (j >= this.Segments.Count)
                            {
                                j = 0;
                            }
                            nextseg = this.Segments[j];
                            break;
                        }

                        segcnt++;
                    }

                    if (string.IsNullOrEmpty(nextseg)) return;

                    while (cnt < max)
                    {
                        if (nextseg.Equals(this.tabControl1.TabPages[cnt].Text))
                        {
                            this.tabControl1.SelectTab(cnt);
                            break;
                        }
                        cnt++;
                    }
                }
                else if (e.KeyCode == Keys.Left)
                {
                    while (segcnt < segmax)
                    {
                        if (segment.Equals(this.Segments[segcnt]))
                        {
                            int j = segcnt - 1;
                            if (j < 0)
                            {
                                j = this.Segments.Count - 1;
                            }
                            nextseg = this.Segments[j];
                            break;
                        }

                        segcnt++;
                    }

                    if( string.IsNullOrEmpty(nextseg)) return;

                    while (cnt < max)
                    {
                        if (nextseg.Equals(this.tabControl1.TabPages[cnt].Text))
                        {
                            this.tabControl1.SelectTab(cnt);
                            break;
                        }
                        cnt++;
                    }
                }
                 */

                e.Handled = true;
                return;
            }
            //this.DataSave();
            if (this.ExitedKeyDown != null)
            {
                this.ExitedKeyDown(e.KeyCode);
            }
            e.Handled = true;
            this.Hide();
        }

        private void frmWordsList_FormClosing(object sender, FormClosingEventArgs e)
        {
            //this.DataSave();
            e.Cancel = true;
            this.Hide();
        }


        public void DataSave()
        {
            FileStream fs = null;
            StreamWriter w = null;

            if (string.IsNullOrEmpty(this.StartupPath))
            {
                return;
            }

            try
            {
                fs = File.Open(this.StartupPath + @"\formsize.txt", FileMode.OpenOrCreate);
                w = new StreamWriter(fs, Encoding.Default);

                w.WriteLine(this.Top.ToString());
                w.WriteLine(this.Left.ToString());
                w.WriteLine(this.Height.ToString());
                w.WriteLine(this.Width.ToString());

                w.Flush();

                //fs.Close();
            }
            catch (Exception excp)
            {
            }
            finally
            {
                if (w != null)
                {
                    w.Close();
                }
                if (fs != null)
                {
                    fs.Close();
                }
            }
        }

        public void DataLoad()
        {
            FileStream fs = null;
            StreamReader r = null;

            if (string.IsNullOrEmpty(this.StartupPath))
            {
                return;
            }

            try
            {
                fs = File.Open(this.StartupPath + @"\formsize.txt", FileMode.OpenOrCreate);
                r = new StreamReader(fs, Encoding.Default);

                string line = "";
                int tmp = 0;
                if ((line = r.ReadLine()) != null)
                {
                    int.TryParse(line, out tmp);
                    if (tmp > 0)
                    {
                        this.Top = tmp;
                    }
                }
                if ((line = r.ReadLine()) != null)
                {
                    int.TryParse(line, out tmp);
                    if (tmp > 0)
                    {
                        this.Left = tmp;
                    }
                }
                if ((line = r.ReadLine()) != null)
                {
                    int.TryParse(line, out tmp);
                    if (tmp > 0)
                    {
                        this.Height = tmp;
                    }
                }
                if ((line = r.ReadLine()) != null)
                {
                    int.TryParse(line, out tmp);
                    if (tmp > 0)
                    {
                        this.Width = tmp;
                    }
                }

                //fs.Close();
            }
            catch (Exception excp)
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
        }

        private void frmWordsList_FormClosed(object sender, FormClosedEventArgs e)
        {
            //this.DataSave();
        }


        private TabPage firstTabPage = null;
        private ListBox firstListBox = null;
        public void TransWords(string word)
        {
            this.firstTabPage = null;
            this.firstListBox = null;

            //YahooAPI（VJE)による漢字変換
            YahooAPI yapi = new YahooAPI();
            string apiid = yapi.GetApiID();
            string word4yahoo = yapi.GetJPWordForyahooAPI(word);


            //ローマ字変換
            List<KeyValue> segments = new List<KeyValue>();
            string urlYahooRoman = yapi.getUrlRoman(word4yahoo, apiid);
            List<WordItem> words = yapi.getWords(urlYahooRoman, out segments);

            // コントロール処理
            this.tabControl1.TabPages.Clear();
            this.Segments = segments;
            StringBuilder sb = new StringBuilder();
            foreach (KeyValue segment in segments)
            {
                sb.Append(segment.Key);
            }
            this.SegmentText = sb.ToString();
            this.rchText.Text = this.SegmentText;

            foreach( KeyValue segment in segments )
            {
                TabPage tab = new TabPage();
                if (firstTabPage == null)
                {
                    firstTabPage = tab;
                }

                tab.Text = segment.Key;
                ListBox list = new ListBox();
                if (firstListBox == null)
                {
                    firstListBox = list;
                }

                foreach (WordItem item in words)
                {
                    if (segment.Key.Equals(item.segment))
                    {
                        list.Items.Add(item);
                    }
                }
                //list.KeyDown += this.lstWords_KeyDown;
                list.KeyUp += this.lstWords_KeyUp;
                list.KeyDown += this.lstWords_KeyDown;
                list.KeyPress += this.lstWords_KeyPress;
                list.Dock = DockStyle.Fill;
                tab.Controls.Add(list);
                this.tabControl1.TabPages.Add(tab);
            }

            // 何も内場合　→　空のタブを作る
            if (this.Segments.Count < 1)
            {
                TabPage tab = new TabPage();
                if (firstTabPage == null)
                {
                    firstTabPage = tab;
                }

                tab.Text = string.Empty;
                ListBox list = new ListBox();
                if (firstListBox == null)
                {
                    firstListBox = list;
                }
                list.Items.Add("");
                list.Dock = DockStyle.Fill;
                tab.Controls.Add(list);
                this.tabControl1.TabPages.Add(tab);
            }

            /*
            this.tabControl1.Focus();
            if (firstTabPage != null)
            {
                firstTabPage.Focus();
            }
            if (firstListBox != null)
            {
                firstListBox.Focus();
            }
             */
        }

        private void frmWordsList_Activated(object sender, EventArgs e)
        {
            this.MoveFocusToDefault();
        }

        /*
        private void SetItems(string segment, List<WordItem> words)
        {
            if( string.IsNullOrEmpty(segment) )
            {
                return;
            }
            this.lstWords.Items.Clear();
            foreach (WordItem item in words)
            {
                if (segment.Equals(item.segment))
                {
                    this.lstWords.Items.Add(item);
                }
            }
        }
         * */

    }
}

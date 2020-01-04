using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using VBIDE = Microsoft.Vbe;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using Extensibility;
using System.Text.RegularExpressions;
using HankakuDeKanjiInput;
using System.IO;
using Microsoft.VisualBasic;

namespace ExcelHankakuDeKanjiInput
{
    // Excelアドインの「発行」（インストーラ作成）の方法
    // 参考URL：
    // https://qiita.com/rhene/items/dd2cf0667f7ca0fd15a8

    public partial class ThisAddIn
    {
        //private VBIDE.VBE app;
        private Microsoft.Vbe.Interop.VBE app;
        private Office.CommandBar cmdBar;
        private Office.CommandBarButton cmdBtn;
        private int startClumn = 0;
        private int endColumn = 0;

        //HankakuDeKanjiInput
        private frmWordsList frmwords = null;//new frmWordsList();
        private string targetWord = string.Empty;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                app = ((Microsoft.Vbe.Interop.VBE)this.Application.VBE);
                cmdBar = app.CommandBars.Add("MyCommandBar", Office.MsoBarPosition.msoBarFloating, false, true);
                cmdBtn = (Office.CommandBarButton)cmdBar.Controls.Add(Office.MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
                cmdBtn.Caption = "My Button";
                cmdBtn.FaceId = 59;
                cmdBtn.Click += new Office._CommandBarButtonEvents_ClickEventHandler(cmdBtn_Click);
                cmdBtn.ShortcutText = "Ctrl+F12";
                cmdBtn.Caption = "半角で漢字変換(&J)";
                cmdBtn.Style = MsoButtonStyle.msoButtonCaption;
                cmdBar.Visible = true;

                /*
                // Menu
                app = ((Microsoft.Vbe.Interop.VBE)this.Application.VBE);
                Office.CommandBar menu = (Office.CommandBar)app.CommandBars.Add("半角で漢字入力", MsoBarPosition.msoBarFloating, false, true);
                menu.Visible = true;
                app.CommandBars["半角で漢字入力"].Visible = true;
                app.CommandBars["半角で漢字入力"].Controls.Add(MsoControlType.msoControlPopup);
                app.CommandBars["半角で漢字入力"].Controls[1].Caption = "変換";
                //app.CommandBars["半角で漢字入力"].Controls[1].OnAction += this.cmdBtn_Click;
                 */

                // HankakuDeKanjiInput
                string baseNasme = Path.GetDirectoryName(this.Application.StartupPath);

                this.frmwords = new frmWordsList()
                {
                    StartupPath = baseNasme
                };
                frmwords.OnWordSelect += this.WordSelected;
                frmwords.ExitedKeyDown += this.ExitedKeyDown;
                frmwords.ExitedKeyPress += this.ExitedKeyPress;

            }
            catch (Exception excp)
            {
                Console.WriteLine(excp.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void cmdBtn_Click(Office.CommandBarButton ctrl, ref bool cancel)
        {
            //try{
            //app.ActiveCodePane.CodeModule.InsertLines(1, "'Hello World!!");


            /*
               //app = ((Microsoft.Vbe.Interop.VBE)ExcelVBEIDE03.Globals.ThisAddIn.Application);
               cmdBar = app.CommandBars.Add("MyCommandBar", Office.MsoBarPosition.msoBarFloating, false, true);
               cmdBtn = (Office.CommandBarButton)cmdBar.Controls.Add(Office.MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
               cmdBtn.Caption = "My Button";
               cmdBtn.FaceId = 59;
               //cmdBtn.Click += new Office._CommandBarButtonEvents_ClickEventHandler(cmdBtn_Click);
               cmdBar.Visible = true;
            */
            var codep = (Microsoft.Vbe.Interop.CodePane)app.CodePanes.Item(1);

            string[] abc = { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t",  "u",  "v",  "w", "x", "y", "z", 
                           };
            int startline = 0;
            int endline = 0;
            int startcolumn = 0;
            int endcolumn = 0;
            codep.GetSelection(out startline, out startcolumn, out endline, out endcolumn);
            //codep.
            var module = codep.CodeModule;
            var line = module.Lines[endline, 1];

            if (endcolumn < 1)
            {
                endcolumn = 1;
            }
            if (line.Length >= endcolumn)
            {
                line = line.Substring(0, endcolumn - 1);
            }
            //MessageBox.Show(line);

            //}catch (Exception){
            //MessageBox.Show("エラーが発生しました。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}

            // HankakuDeKanjiInput
            string word = line;
            try
            {

                Regex reg = new Regex("[-a-zA-Z0-9]+");
                MatchCollection matchCol = reg.Matches(word);
                if (matchCol.Count > 0)
                {
                    this.targetWord = matchCol[matchCol.Count - 1].Value;
                }
                else
                {
                    return;
                }


                //Shift JISとして文字列に変換
                byte[] bytesData = null;
                //UTF-8として変換
                bytesData = System.Text.Encoding.UTF8.GetBytes(line);
                string endcol_sjis = "";
                endcol_sjis = System.Text.Encoding.GetEncoding(932).GetString(bytesData);
                Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
                int line_length = sjisEnc.GetByteCount(endcol_sjis);


                this.startClumn = line.IndexOf(this.targetWord);
                this.endColumn = line.LastIndexOf(this.targetWord) + this.targetWord.Length;
                if (this.startClumn < 0)
                {
                    this.startClumn = 0;
                }
                if (this.endColumn < 0)
                {
                    this.endColumn = 0;
                }
                //this.startClumn = endcol_sjis.IndexOf(this.targetWord);
                //this.endColumn = line_length - this.startClumn;

                //frmwords.TransWords(word);
                frmwords.TransWords(this.targetWord);
                frmwords.Show();
                //frmwords.Left = p.X;
                //frmwords.Top = p.Y;
                frmwords.Focus();
            }
            catch (Exception excp)
            {
                Console.WriteLine(excp.Message);
            }
        }
        private void WordSelected(string word)
        {


            word = word.Replace("\r", "").Replace("\n", "");
            this.InsertString(word);
            this.targetWord = string.Empty;
        }
        private void ExitedKeyDown(Keys key)
        {
            //if (key.Equals(Keys.Control))
            //{
                this.InsertString(this.frmwords.Words);
            //}
        }
        private void ExitedKeyPress(char keyChar)
        {
            // 何もしない
            //this.InsertString(keyChar.ToString());


            /* // this.InsertString()　へ移行
            var codep = (Microsoft.Vbe.Interop.CodePane)app.CodePanes.Item(1);


            int startline = 0;
            int endline = 0;
            int startcolumn = 0;
            int endcolumn = 0;
            codep.GetSelection(out startline, out startcolumn, out endline, out endcolumn);

          
            
            var module = codep.CodeModule;
            var line = module.Lines[endline, 1];
            string afterLine = string.Empty;
            string beforeLine = string.Empty;
            string mainLine = string.Empty;

            if (endcolumn < 1)
            {
                endcolumn = 1;
            }
            if (line.Length >= endcolumn)
            {
                beforeLine = line.Substring(0, startcolumn - 1);
                mainLine = line.Substring(startcolumn, endcolumn - startcolumn);
                afterLine = line.Substring(endcolumn, line.Length - endcolumn);
            }

            module.DeleteLines(endline);
            module.InsertLines(endline, beforeLine + mainLine + afterLine);
            */
        }
        private void InsertString(string str)
        {
            try
            {
                var codep = (Microsoft.Vbe.Interop.CodePane)app.CodePanes.Item(1);


                int startline = 0;
                int endline = 0;
                int startcolumn = 0;
                int endcolumn = 0;
                codep.GetSelection(out startline, out startcolumn, out endline, out endcolumn);



                var module = codep.CodeModule;
                var line = module.Lines[endline, 1];
                string afterLine = string.Empty;
                string beforeLine = string.Empty;
                string mainLine = string.Empty;

                if (endcolumn < 1)
                {
                    endcolumn = 1;
                }

                beforeLine = line.Substring(0, this.startClumn);
                afterLine = line.Substring(this.endColumn);

                //Shift JISとして文字列に変換
                byte[] bytesData = null;
                //UTF-8として変換
                bytesData = System.Text.Encoding.UTF8.GetBytes(beforeLine + str);
                string endcol_sjis = "";
                endcol_sjis = System.Text.Encoding.GetEncoding(932).GetString(bytesData);
                Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
                int num = sjisEnc.GetByteCount(endcol_sjis);

                module.DeleteLines(endline);
                module.InsertLines(endline, beforeLine + str + afterLine);
                codep.SetSelection(endline, num * 2, endline, num * 2);
            }
            catch (Exception excp)
            {
                Console.WriteLine(excp.Message);
            }
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }


    //    /*
    //     * 参考URL：http://www.ka-net.org/blog/?p=5552
    //     */
    //    //Guidは要変更
    //[ComVisible(true), Guid("D6FCC113-0156-494C-824B-AE73D83DD78B"), ProgId("ExcelHankakuDeKanjiInput.Connect")]
    //public class Connect : Object, Extensibility.IDTExtensibility2
    //{
    //    //private VBIDE.VBE app;
    //    private Microsoft.Vbe.Interop.VBE app;
    //    private Office.CommandBar cmdBar;
    //    private Office.CommandBarButton cmdBtn;
    //    public Connect() { }

    //    public void OnConnection(object application, ext_ConnectMode ConnectMode, object AddInInst, ref System.Array custom)
    //    {
    //        app = ((Microsoft.Vbe.Interop.VBE)application);
    //        cmdBar = app.CommandBars.Add("MyCommandBar", Office.MsoBarPosition.msoBarFloating, false, true);
    //        cmdBtn = (Office.CommandBarButton)cmdBar.Controls.Add(Office.MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
    //        cmdBtn.Caption = "My Button";
    //        cmdBtn.FaceId = 59;
    //        cmdBtn.Click += new Office._CommandBarButtonEvents_ClickEventHandler(cmdBtn_Click);
    //        cmdBar.Visible = true;
    //        //cmdBar.accKeyboardShortcut 
    //        //cmdBtn.accKeyboardShortcut

    //        // https://msdn.microsoft.com/ja-jp/library/office/ff864972.aspx
    //        // http://stackoverflow.com/questions/14341138/keyboard-shortcuts-in-excel-menu-commandbar
    //        // cmdBtn.ShortcutText

    //        // 独自メニュー割り当て
    //        //http://officetanaka.net/excel/vba/tips/tips05.htm

    //        /*
    //        app = ((Microsoft.Vbe.Interop.VBE)ExcelVBEIDE03.Globals.ThisAddIn.Application.VBE);
    //        Office.CommandBar menu = app.CommandBars.Add("TestMenu", missing, true, true);
    //        app.CommandBars["TestMenu"].Visible = true;          
    //        app.CommandBars["TestMenu"].Controls.Add(MsoControlType.msoControlPopup);
    //        app.CommandBars["TestMenu"].Controls[0].Caption = "新規メニュー";
    //        //app.CommandBars["TestMenu"].Controls[0].OnAction += 
    //        */

    //        /*
    //            With Application.CommandBars("New_Bar")
    //    .Visible = True
    //    .Controls.Add Type:=msoControlPopup
    //    .Controls(1).Caption = "新規メニュー"
    //    With .Controls(1)
    //        .Controls.Add Type:=msoControlButton
    //        With .Controls(1)
    //            .Caption = "追加コマンド１"
    //            .OnAction = "Msg_1"
    //        End With

    //        .Controls.Add Type:=msoControlPopup
    //        With .Controls(2)
    //            .Caption = "追加コマンド２"
    //            .Controls.Add Type:=msoControlButton
    //            .Controls(1).Caption = "サブコマンド"
    //            .Controls(1).OnAction = "Msg_2"
    //        End With
    //          */

    //        /*
    //        //Vbe.Interop.Windows
    //        app = ((Microsoft.Vbe.Interop.VBE)ExcelHankakuDeKanjiInput.Globals.ThisAddIn.Application);
    //        cmdBar = app.CommandBars.Add("MyCommandBar", Office.MsoBarPosition.msoBarFloating, false, true);
    //        cmdBtn = (Office.CommandBarButton)cmdBar.Controls.Add(Office.MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
    //        cmdBtn.Caption = "My Button";
    //        cmdBtn.FaceId = 59;
    //        cmdBtn.Click += new Office._CommandBarButtonEvents_ClickEventHandler(cmdBtn_Click);
    //        cmdBar.Visible = true;
    //        cmdBar.ShowPopup();
    //        */
    //        /*

    //        var codep = (Microsoft.Vbe.Interop.CodePane)app.CodePanes.Item(0);

    //        string[] abc = { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t",  "u",  "v",  "w", "x", "y", "z", 
    //                       };
    //        int startline = 0;
    //        int endline = 0;
    //        int startcolumn = 0;
    //        int endcolumn =0;
    //        codep.GetSelection(out startline, out startcolumn, out endline,  out endcolumn);
    //        //codep.

    //        var module = codep.CodeModule;
    //        var line = module.Lines[endline, 1];

    //        if (endcolumn < 1)
    //        {
    //            endcolumn = 1;
    //        }
    //        if (line.Length >= endcolumn)
    //        {
    //            line = line.Substring(0, endcolumn - 1);
    //        }
    //        //MessageBox.Show(line);
    //         */
    //    }
    //    public void cmdBtn_Click(Office.CommandBarButton ctrl, ref bool cancel)
    //    {
    //        //try{
    //        //app.ActiveCodePane.CodeModule.InsertLines(1, "'Hello World!!");


    //        /*
    //           //app = ((Microsoft.Vbe.Interop.VBE)ExcelVBEIDE03.Globals.ThisAddIn.Application);
    //           cmdBar = app.CommandBars.Add("MyCommandBar", Office.MsoBarPosition.msoBarFloating, false, true);
    //           cmdBtn = (Office.CommandBarButton)cmdBar.Controls.Add(Office.MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
    //           cmdBtn.Caption = "My Button";
    //           cmdBtn.FaceId = 59;
    //           //cmdBtn.Click += new Office._CommandBarButtonEvents_ClickEventHandler(cmdBtn_Click);
    //           cmdBar.Visible = true;
    //        */
    //        var codep = (Microsoft.Vbe.Interop.CodePane)app.CodePanes.Item(1);

    //        string[] abc = { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t",  "u",  "v",  "w", "x", "y", "z", 
    //                       };
    //        int startline = 0;
    //        int endline = 0;
    //        int startcolumn = 0;
    //        int endcolumn = 0;
    //        codep.GetSelection(out startline, out startcolumn, out endline, out endcolumn);
    //        //codep.

    //        var module = codep.CodeModule;
    //        var line = module.Lines[endline, 1];

    //        if (endcolumn < 1)
    //        {
    //            endcolumn = 1;
    //        }
    //        if (line.Length >= endcolumn)
    //        {
    //            line = line.Substring(0, endcolumn - 1);
    //        }
    //        MessageBox.Show(line);

    //        //}catch (Exception){
    //        //MessageBox.Show("エラーが発生しました。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
    //        //}
    //    }

    //    public void OnDisconnection(ext_DisconnectMode RemoveMode, ref System.Array custom)
    //    {
    //        if (cmdBtn != null)
    //        {
    //            Marshal.ReleaseComObject(cmdBtn);
    //            cmdBtn = null;
    //        }
    //        if (cmdBar != null)
    //        {
    //            cmdBar.Delete();
    //            Marshal.ReleaseComObject(cmdBar);
    //            cmdBar = null;
    //        }
    //        if (app != null)
    //        {
    //            Marshal.ReleaseComObject(app);
    //            app = null;
    //        }
    //        GC.Collect();
    //        GC.WaitForPendingFinalizers();
    //    }
    //    public void OnAddInsUpdate(ref System.Array custom) { }
    //    public void OnStartupComplete(ref System.Array custom) { }
    //    public void OnBeginShutdown(ref System.Array custom) { }



    //}
}

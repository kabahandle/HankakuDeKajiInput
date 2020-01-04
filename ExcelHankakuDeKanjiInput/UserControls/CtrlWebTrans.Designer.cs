namespace MyVSNetAddin2
{
    partial class CtrlWebTrans
    {
        /// <summary> 
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary> 
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を 
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.webTrans = new System.Windows.Forms.WebBrowser();
            this.SuspendLayout();
            // 
            // webTrans
            // 
            this.webTrans.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webTrans.Location = new System.Drawing.Point(0, 0);
            this.webTrans.MinimumSize = new System.Drawing.Size(20, 20);
            this.webTrans.Name = "webTrans";
            this.webTrans.ScriptErrorsSuppressed = true;
            this.webTrans.Size = new System.Drawing.Size(150, 150);
            this.webTrans.TabIndex = 0;
            // 
            // CtrlWebTrans
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.webTrans);
            this.Name = "CtrlWebTrans";
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.WebBrowser webTrans;

    }
}

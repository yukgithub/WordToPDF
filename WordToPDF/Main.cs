using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using RefWord = Microsoft.Office.Interop.Word;

namespace WordToPDF
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 入力ボタンクリック処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnInput_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "フォルダを選択";

                if(dialog.ShowDialog(this) == DialogResult.OK)
                {
                    this.txtInput.Text = dialog.SelectedPath;
                    this.txtOutput.Text = dialog.SelectedPath;
                }


            }
        }

        /// <summary>
        /// 出力ボタンクリック処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOutput_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "フォルダを選択";

                if (dialog.ShowDialog(this) == DialogResult.OK)
                    this.txtOutput.Text = dialog.SelectedPath;


            }

        }

        /// <summary>
        /// 変換ボタンクリック処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConvert_Click(object sender, EventArgs e)
        {

            RefWord.Application word = new RefWord.Application();
            RefWord.Documents docs = word.Documents;
            RefWord.Document doc = null;

            this.toolComplete.Text = string.Empty;
            this.Cursor = Cursors.WaitCursor;

            try
            {

                string[] files = Directory.GetFiles(this.txtInput.Text, "*.doc*");

                foreach (string f in files)
                {
                    this.toolComplete.Text = Path.GetFileName(f);

                    var attribute = File.GetAttributes(f);
                    var fn = Path.GetFileNameWithoutExtension(f);

                    //隠しファイル（テンポラリ）は飛ばす
                    if ((attribute & FileAttributes.Hidden) == FileAttributes.Hidden)
                        continue;

                    //出力
                    doc = docs.Open(f, ReadOnly:true);
                    doc.ExportAsFixedFormat(this.txtOutput.Text + "\\" + fn + ".pdf", RefWord.WdExportFormat.wdExportFormatPDF);
                    doc.Close();

                }

                this.toolComplete.Text = "完了";

            } catch(Exception ex)
            {

                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            } finally
            {

                MessageBox.Show(GC.GetTotalMemory(false).ToString());

                
                //GC.Collect();
               // GC.WaitForPendingFinalizers();
                //GC.Collect();
                

                if (word != null) word.Quit();

                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(docs);
                Marshal.ReleaseComObject(word);

                
                //GC.Collect();
                //GC.WaitForPendingFinalizers();
                //GC.Collect();
                

                MessageBox.Show(GC.GetTotalMemory(false).ToString());

            }

            this.Cursor = Cursors.Default;

        }

        /// <summary>
        /// 起動時処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmMain_Load(object sender, EventArgs e)
        {
            this.txtInput.Text = Directory.GetCurrentDirectory();
            this.txtOutput.Text = Directory.GetCurrentDirectory();
        }
    }
}

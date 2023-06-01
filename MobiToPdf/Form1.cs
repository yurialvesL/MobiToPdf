using System;
using System.Windows.Forms;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Properties;
using System.Runtime.InteropServices;

namespace MobiToPdf
{
    public partial class frm_mobiPdf : Form
    {
        public frm_mobiPdf()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            lbl_title = new Label();
            txt_mobi = new TextBox();
            btn_mobi = new Button();
            progressBar1 = new ProgressBar();
            btn_convert = new Button();
            lbl_pdf = new Label();
            SuspendLayout();
            // 
            // lbl_title
            // 
            lbl_title.AutoSize = true;
            lbl_title.Location = new Point(39, 32);
            lbl_title.Name = "lbl_title";
            lbl_title.Size = new Size(80, 15);
            lbl_title.TabIndex = 0;
            lbl_title.Text = "Arquivo Mobi";
            // 
            // txt_mobi
            // 
            txt_mobi.Enabled = false;
            txt_mobi.Location = new Point(39, 50);
            txt_mobi.Name = "txt_mobi";
            txt_mobi.Size = new Size(329, 23);
            txt_mobi.TabIndex = 1;
            // 
            // btn_mobi
            // 
            btn_mobi.Location = new Point(144, 79);
            btn_mobi.Name = "btn_mobi";
            btn_mobi.Size = new Size(100, 23);
            btn_mobi.TabIndex = 2;
            btn_mobi.Text = "Carregar Mobi";
            btn_mobi.UseVisualStyleBackColor = true;
            btn_mobi.Click += btn_mobi_Click;
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(39, 204);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(329, 23);
            progressBar1.TabIndex = 3;
            // 
            // btn_convert
            // 
            btn_convert.Location = new Point(157, 132);
            btn_convert.Name = "btn_convert";
            btn_convert.Size = new Size(75, 23);
            btn_convert.TabIndex = 4;
            btn_convert.Text = "Converter";
            btn_convert.UseVisualStyleBackColor = true;
            btn_convert.Click += btn_convert_Click;
            // 
            // lbl_pdf
            // 
            lbl_pdf.AutoSize = true;
            lbl_pdf.Location = new Point(175, 176);
            lbl_pdf.Name = "lbl_pdf";
            lbl_pdf.Size = new Size(0, 15);
            lbl_pdf.TabIndex = 5;
            // 
            // frm_mobiPdf
            // 
            ClientSize = new Size(412, 266);
            Controls.Add(lbl_pdf);
            Controls.Add(btn_convert);
            Controls.Add(progressBar1);
            Controls.Add(btn_mobi);
            Controls.Add(txt_mobi);
            Controls.Add(lbl_title);
            Name = "frm_mobiPdf";
            Text = "MobiToPDF";
            ResumeLayout(false);
            PerformLayout();
        }

        private Label lbl_title;
        private TextBox txt_mobi;
        private ProgressBar progressBar1;
        private Button btn_convert;
        private Label lbl_pdf;
        private Button btn_mobi;

        private void btn_mobi_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog op = new OpenFileDialog())
            {
                op.InitialDirectory = "C:\\";
                op.Title = "Selecione o Mobi";
                op.Filter = "Mobi Files (*.mobi)|";
                op.FilterIndex = 0;
                op.RestoreDirectory = true;

                if (op.ShowDialog() == DialogResult.OK)
                {
                    string caminho = op.FileName;
                    txt_mobi.Text = caminho;
                }
            }
        }

        private void btn_convert_Click(object sender, EventArgs e)
        {
            Document doc = new Document(txt_mobi.Text);
            string caminho = "";
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.PageSetup.Margins = Aspose.Words.Margins.Mirrored;
            builder.PageSetup.PageHeight = 1200;
            builder.PageSetup.PageWidth = 900;
            builder.PageSetup.Orientation = Aspose.Words.Orientation.Portrait;

            using (OpenFileDialog op = new OpenFileDialog())
            {
                op.InitialDirectory = "C:\\";
                op.Title = "Salvar arquivo";
                op.Filter = "PDF Files (*.pdf)|";
                op.FilterIndex = 0;
                op.RestoreDirectory = true;
                op.CheckFileExists = false;
                op.ShowReadOnly = true;
                if (op.ShowDialog() == DialogResult.OK)
                {
                    caminho = op.FileName;

                }
                try
                {
                    doc.Save(caminho,SaveFormat.Pdf);
                    

                }
                catch (Exception ex)
                {

                    MessageBox.Show($"s{ex.Message}", "Error Generate", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
using System;
using System.Windows.Forms;
using System.Globalization;
using Aspose.Words;
using System.Threading.Tasks;
using Aspose.Words.Properties;
using System.Runtime.InteropServices;
using System.Diagnostics;

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
            Btn_Cancel = new Button();
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
            progressBar1.Location = new Point(39, 194);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(329, 23);
            progressBar1.TabIndex = 3;
            // 
            // btn_convert
            // 
            btn_convert.Location = new Point(109, 135);
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
            // Btn_Cancel
            // 
            Btn_Cancel.Location = new Point(190, 135);
            Btn_Cancel.Name = "Btn_Cancel";
            Btn_Cancel.Size = new Size(75, 23);
            Btn_Cancel.TabIndex = 6;
            Btn_Cancel.Text = "Cancelar";
            Btn_Cancel.UseVisualStyleBackColor = true;
            Btn_Cancel.Click += Btn_Cancel_Click;
            // 
            // frm_mobiPdf
            // 
            ClientSize = new Size(412, 258);
            Controls.Add(Btn_Cancel);
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
        private Button Btn_Cancel;
        private CancellationTokenSource cancellationTokenSource;
        private int progress = 1;
        private string caminho = "";

        private void btn_mobi_Click(object sender, EventArgs e)
        {
            bool validate = false;
            using (OpenFileDialog op = new OpenFileDialog())
            {
                op.InitialDirectory = "C:\\";
                op.Title = "Selecione o Mobi";
                op.Filter = "Mobi Files (*.mobi)|";
                op.FilterIndex = 0;
                op.RestoreDirectory = true;


                while (validate != true)
                {
                    if(op.ShowDialog() == DialogResult.OK)
                    {
                         caminho = op.FileName;
                        if (!op.FileName.EndsWith(".mobi"))
                        {
                            MessageBox.Show("O arquivo selecionado não é um arquivo mobi\nEscolha o arquivo novamente", "Arquivo incorreto", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else 
                        { 
                            validate = true; 
                        }
                       

                    }
                   
                }
                txt_mobi.Text = caminho;

                btn_mobi.Enabled = false;
            }
        }

        private async void btn_convert_Click(object sender, EventArgs e)
        {
            Document doc = new Document(txt_mobi.Text);
            cancellationTokenSource = new CancellationTokenSource();
            bool complete = false;
            Task tk =   Task.WhenAll(ToPdf(doc, cancellationTokenSource));
            while (await Task.WhenAny(tk, Task.Delay(5000)) != tk)
            {
                
                    progressBar1.Value += 5;
                

            }
            progressBar1.Value = progressBar1.Maximum;
            Clearall();



        }
        private async Task ToPdf(Document doc, CancellationTokenSource token)
        {
            string caminho = "";
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.PageSetup.Margins = Aspose.Words.Margins.Mirrored;
            builder.PageSetup.PageHeight = 1200;
            builder.PageSetup.PageWidth = 900;
            builder.PageSetup.Orientation = Aspose.Words.Orientation.Portrait;
            builder.PageSetup.Orientation = Aspose.Words.Orientation.Portrait;
            btn_convert.Enabled = false;

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
                    caminho = op.FileName + ".pdf";

                }
                try
                {
                    await Task.Run(() => { doc.Save(caminho, SaveFormat.Pdf); });                        // método que converte

                }
                catch (Exception ex)
                {

                    MessageBox.Show($"s{ex.Message}", "Error Generate", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }
        private void  Clearall()
        {

            if (progressBar1.Value == 100)
            {
                MessageBox.Show("Download concluido com sucesso", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                progressBar1.Value = 0;
                progress = 0;
                btn_convert.Enabled = true;
                btn_mobi.Enabled = true;
                txt_mobi.Text = "";
            }
        }

        private void Btn_Cancel_Click(object sender, EventArgs e)
        {
            cancellationTokenSource.Cancel();
        }
    }
}
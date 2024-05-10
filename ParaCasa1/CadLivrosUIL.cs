using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;


namespace ParaCasa1
{
    public partial class CadLivrosUIL : Form
    {
        public CadLivrosUIL()
        {
            InitializeComponent();
        }

        private void CadLivrosUIL_Load(object sender, EventArgs e)
        {
            LivroBLL.conecta();
            if (Erro.getErro())
            {
                MessageBox.Show(Erro.getMsg());
                System.Windows.Forms.Application.Exit();
            }
        }

        private void CadLivrosUIL_FormClosed(object sender, FormClosedEventArgs e)
        {
            LivroBLL.desconecta();
        }

         private void button1_Click(object sender, EventArgs e)
        {
             LivroBLL.getProximo();
             while (!Erro.getErro())
             {
                 listBox1.Items.Add("Titulo = " + Livro.getTitulo() + " escrito por " + Livro.getAutor());
                 LivroBLL.getProximo();
             }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            Document doc = word.Documents.Add();
            Selection texto = word.Selection;
            doc.Activate();
            LivroBLL.getProximo();
            while (!Erro.getErro())
            {
                texto.TypeText("Titulo = " + Livro.getTitulo() + " escrito por " + Livro.getAutor()+ "\n");
                texto.TypeParagraph();
                LivroBLL.getProximo();
            }
            doc.SaveAs(@"c:\Users\unisanta\Desktop\Word\listagem.docx");
            doc.SaveAs(@"c:\Users\unisanta\Desktop\Word\listagem.pdf", WdSaveFormat.wdFormatPDF);
            doc.Close();
            word.Quit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Workbooks.Add();
            Worksheet ws = excel.ActiveSheet;
            int lin = 1;
            LivroBLL.getProximo();
            while (!Erro.getErro())
            {
                ws.Cells[lin, 1] = Livro.getTitulo() + " escrito por ";
                ws.Cells[lin, 2] = Livro.getAutor();
                ++lin;
                LivroBLL.getProximo();
            }

            ws.SaveAs(@"c:\Users\unisanta\Desktop\Word\listagem2.xlsx");
            excel.Quit();
        }
    }
}

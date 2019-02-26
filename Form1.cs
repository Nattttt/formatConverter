using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Slides;
using Aspose.Slides.Export;
using Spire.Doc;
using Spire.Xls;
using Spire.Doc.Documents;
using System.IO;

namespace convertidorFormatos
{
    public partial class Form1 : Form
    {

        string ruta;
        string rutaout;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog op1 = new OpenFileDialog();

            if (op1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ruta = op1.FileName;
                label2.Text = Path.GetFileName(ruta).ToString();
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            
            SaveFileDialog op2 = new SaveFileDialog();

            if (op2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                rutaout = op2.FileName;
                label1.Text = Path.GetFileName(rutaout).ToString();
            }


        }

        private void btnRun_Click(object sender, EventArgs e)
        {



            // The path to the documents directory.
            

            // Open the ODP file
           // Presentation pres = new Presentation(ruta);
            
            // Saving the ODP presentation to PPTX format
           // pres.Save(rutaout + "AccessOpenDoc_out.pptx", SaveFormat.Pptx);


            

             
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(ruta);
            workbook.SaveToFile(rutaout +"Result.ods", Spire.Xls.FileFormat.ODS);
           
           /*  Document doc = new Document();
             doc.LoadFromFile(ruta);
             doc.SaveToFile(rutaout +"output.odt", Spire.Doc.FileFormat.Odt);

            /*
            Document doc = new Document();
            doc.LoadFromFile("SampleODTFile.odt");
            doc.SaveToFile(rutaout +"output.docx", FileFormat.Docx);
            
             //convertir excel

             Workbook workbook = new Workbook();
            workbook.LoadFromFile("Sample.xlsx");
            workbook.SaveToFile("Result.ods",FileFormat.ODS);
             
             */

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {
           

        }

     
    }
}

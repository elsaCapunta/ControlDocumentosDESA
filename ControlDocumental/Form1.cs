using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ControlDocumental.Properties;
using System.IO;

 
namespace ControlDocumental
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string fileNAme = "C:\\BDCD\\bdCD.mdb";
            if (!existDB(fileNAme))
            {
                bool realiza = CreateNewAccessDatabase(fileNAme);
                if(realiza){
                    MessageBox.Show("Se ha creado la BD con éxito");
                }
                else
                {
                    MessageBox.Show("Ha ocurrido un error al crear la BD");
                }
            }
            else
            {
                MessageBox.Show("La BD ya existe");
            }
            
        }

       

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        

        public bool existDB(string curFile)
        {
            return File.Exists(curFile);
        }

        public bool CreateNewAccessDatabase(string fileName)
        {
            bool result = false;

            ADOX.Catalog cat = new ADOX.Catalog();
            ADOX.Table tbl_OC = new ADOX.Table();
            ADOX.Table tbl_EP = new ADOX.Table();
            ADOX.Table tbl_pesoDoc = new ADOX.Table();
            ADOX.Table tbl_tipoProd = new ADOX.Table();
            ADOX.Table tbl_tipoDoc = new ADOX.Table();

            //CREA TABLA tbl_OC
            tbl_OC.Name = "tbl_OC";
            tbl_OC.Columns.Append("OC", ADOX.DataTypeEnum.adVarWChar, 15);
            tbl_OC.Columns.Append("F_Ini", ADOX.DataTypeEnum.adVarWChar, 10);
            tbl_OC.Columns.Append("F_Fin", ADOX.DataTypeEnum.adVarWChar, 10);
            tbl_OC.Columns.Append("fecha", ADOX.DataTypeEnum.adDBDate, 50);

            //CREA TABLA tbl_EP
            tbl_EP.Name = "tbl_EP";
            tbl_EP.Columns.Append("id_OC", ADOX.DataTypeEnum.adNumeric, 10);
            tbl_EP.Columns.Append("F_Ini", ADOX.DataTypeEnum.adVarWChar, 10);
            tbl_EP.Columns.Append("F_Fin", ADOX.DataTypeEnum.adVarWChar, 10);
            tbl_EP.Columns.Append("fecha", ADOX.DataTypeEnum.adDBDate, 50);
            tbl_EP.Columns.Append("hora", ADOX.DataTypeEnum.adDBTime, 50);

            //CREA TABLA tbl_tipoDoc
            tbl_tipoDoc.Name = "tbl_tipoDoc";
            tbl_tipoDoc.Columns.Append("producto", ADOX.DataTypeEnum.adVarWChar, 10);
            
            
            //CREA TABLA tbl_pesoDoc
            tbl_pesoDoc.Name = "tbl_pesoDoc";
            tbl_pesoDoc.Columns.Append("id_form", ADOX.DataTypeEnum.adNumeric, 10);
            tbl_pesoDoc.Columns.Append("cantidad", ADOX.DataTypeEnum.adNumeric, 10);
            tbl_pesoDoc.Columns.Append("peso", ADOX.DataTypeEnum.adDouble, 5);


            try
            {
                cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fileName + "; Jet OLEDB:Engine Type=5");
                cat.Tables.Append(tbl_OC);
                cat.Tables.Append(tbl_EP);
                cat.Tables.Append(tbl_pesoDoc);
                //Now Close the database
                ADODB.Connection con = cat.ActiveConnection as ADODB.Connection;
                if (con != null)
                    con.Close();
                result = true;
            }
            catch (Exception ex)
            {
                result = false;
            }
            cat = null;
            return result;
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }
    }
}

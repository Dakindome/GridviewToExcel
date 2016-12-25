using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Data.SqlClient;
using System.Configuration;
using Microsoft.Office.Interop.Excel;





namespace GridviewToExcel
{
    public partial class Form1 : Form
    {
        public string MyConnectionStrings = ConfigurationManager.ConnectionStrings["SQLconnection"].ConnectionString;
 
        public Form1()
        {
            InitializeComponent();
        }

        private void LoadGrid_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
           
            SqlConnection con = new SqlConnection(MyConnectionStrings);
           
            string sql = "SELECT id ,state ,address,postcode,latitude,longitude  FROM stations where postcode like '%74%'";
            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            da.Fill(ds, "AAA");
            dataGridView1.DataSource = ds.Tables["AAA"] ;

            label3.Text = datefrom.Value.ToString("dd-MMM-yyyy");
            label4.Text = dateto.Value.ToString("dd-MMM-yyyy");

        }

        private void Ex2Excel_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = Excel.Workbooks.Add(XlSheetType.xlWorksheet);

            Worksheet ws = (Worksheet)Excel.ActiveSheet;
            Excel.Visible = true;

            ws.Cells[1, 1] = "id";
            ws.Cells[1, 2] = "state";
            ws.Cells[1, 3] = "address";
            ws.Cells[1, 4] = "postcode";
            ws.Cells[1, 5] = "latitude";
            ws.Cells[1, 6] = "longitude";

            for (int j = 2; j <= dataGridView1.Rows.Count; j++)
            {
                for (int i = 1; i <= 6; i++)
                {
                    ws.Cells[j,i] =dataGridView1.Rows[j-2].Cells[i-1].Value;
                }
            }

        }
    }
}

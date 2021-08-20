using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    public static string path = @"C:\Users\iamma\Desktop\ExcelToSql\ExcelFile\medicine.xlsx";
    public static string conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+ path + "; Extended Properties=Excel 12.0";
    DataSet ds;
    protected void BtnUploadFile_Click(object sender, EventArgs e)
    {
        OleDbConnection OleDbCon = new OleDbConnection(conStr);
        OleDbCon.Open();
        OleDbCommand cmd = new OleDbCommand("SELECT ProductType, ProductName, Category, SubCategory, Tags, Price FROM [Sheet1$]", OleDbCon);
        OleDbDataAdapter adaptr = new OleDbDataAdapter(cmd);

        ds = new DataSet();
        adaptr.Fill(ds);

        OleDbDataReader dr = cmd.ExecuteReader();

        string connectstr = @"Data Source=(LocalDB)\MSSQLLocalDb;Initial Catalog=CrystalRepTest;Integrated Security=True" ;

        SqlBulkCopy sbc = new SqlBulkCopy(connectstr);
        sbc.DestinationTableName = "ItemInfo";
        sbc.WriteToServer(dr);
        OleDbCon.Close();
    }
}
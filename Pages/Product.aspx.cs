using System;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace ScaleModelsExcel.Pages
{
    public partial class Product : System.Web.UI.Page
    {
        private static string ConStr = string.Empty;
        private static readonly string fileName = "modelos.xls";
        private static string path = string.Empty;
        private static string ext = string.Empty;
        private static string id = string.Empty;
        readonly DataTable DtExcel = new DataTable();
        protected void Page_Load(object sender, EventArgs e)
        {
            if(!IsPostBack)
            {
                if (!string.IsNullOrWhiteSpace(Request.QueryString["id"])) {
                    id = Request.QueryString["id"];
                    ConnectToExcel();
                    if (DtExcel.Rows.Count > 0)
                    {
                        FillPage();
                    }
                    else
                    {

                    }
                }
            }
        }

        private void ConnectToExcel()
        {
            //getting the path of the file   
            path = Server.MapPath("~/db/" + fileName); // string path = Server.MapPath("~/db/" + FileUpload1.FileName);
                                                       //saving the file inside the MyFolder of the server  
                                                       //FileUpload1.SaveAs(path);
                                                       //checking whether extension is .xls or .xlsx  
                                                       //Extension of the file upload control saving into ext because   
                                                       //there are two types of extension .xls and .xlsx of Excel   
            ext = Path.GetExtension(fileName).ToLower(); //  string ext = Path.GetExtension(FileUpload1.FileName).ToLower();

            if (ext.Trim() == ".xls")
            {
                //connection string for that file which extension is .xls  
                ConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
            }
            else if (ext.Trim() == ".xlsx")
            {
                //connection string for that file which extantion is .xlsx  
                ConStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            }

            LitStatus.Text = "";

            using (OleDbConnection con = new OleDbConnection(ConStr))
            {
                try
                {
                    con.Open();

                    string query = "select top 1 * from[Sheet1$] where Scale18 = '" + id + "';" ;
                    OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(query, con);
                    oleDbDataAdapter.Fill(DtExcel);
                }
                catch (Exception ex)
                {
                    LitStatus.Text = "Error: " + ex.ToString();
                }
                finally
                {
                    if(con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                    con.Dispose();
                }
            }
        }

        private void FillPage()
        {
            Label1.Text = DtExcel.Rows[0]["Collection"].ToString();
            Label2.Text = DtExcel.Rows[0]["Brand"].ToString();
            Label3.Text = DtExcel.Rows[0]["Model"].ToString();
            Label4.Text = DtExcel.Rows[0]["CarYear"].ToString();
            Label5.Text = DtExcel.Rows[0]["Maker"].ToString();
            Label6.Text = DtExcel.Rows[0]["Scale"].ToString();
            Label7.Text = DtExcel.Rows[0]["PartNumber"].ToString();
            Label8.Text = DtExcel.Rows[0]["CarNumber"].ToString();
            Label9.Text = DtExcel.Rows[0]["ColourSponsor"].ToString();
            Label10.Text = DtExcel.Rows[0]["Driver"].ToString();
            TextBox1.Text = DtExcel.Rows[0]["Details"].ToString();
            TextBox1.Enabled = false;
            Label11.Text = DtExcel.Rows[0]["ModelDate"].ToString();
            Label12.Text = DtExcel.Rows[0]["Serial"].ToString();
            Label13.Text = DtExcel.Rows[0]["Ledition"].ToString();
            TextBox2.Text = DtExcel.Rows[0]["Comments"].ToString();
            TextBox2.Enabled = false;

            if (!String.IsNullOrEmpty(id))
            {
                if (Directory.Exists(Server.MapPath("~/Images/" + id + "/"))) 
                {
                    GetImages();
                }
                else
                {
                    Image1.ImageUrl = "~/Images/nophoto.jpg";
                    Image1.AlternateText = Image1.ImageUrl;
                    Image2.ImageUrl = "~/Images/nophoto.jpg";
                    Image2.AlternateText = Image2.ImageUrl;
                    Image3.ImageUrl = "~/Images/nophoto.jpg";
                    Image3.AlternateText = Image3.ImageUrl;
                    Image4.ImageUrl = "~/Images/nophoto.jpg";
                    Image4.AlternateText = Image4.ImageUrl;
                }
            }
            else
            {
                Image1.ImageUrl = "~/Images/nophoto.jpg";
                Image1.AlternateText = Image1.ImageUrl;
                Image2.ImageUrl = "~/Images/nophoto.jpg";
                Image2.AlternateText = Image2.ImageUrl;
                Image3.ImageUrl = "~/Images/nophoto.jpg";
                Image3.AlternateText = Image3.ImageUrl;
                Image4.ImageUrl = "~/Images/nophoto.jpg";
                Image4.AlternateText = Image4.ImageUrl;
            }
        }

        private void GetImages()
        {
            string[] images = Directory.GetFiles(Server.MapPath("~/Images/" + id + "/"));

            int i = 0;

            foreach(string image in images)
            {
                string imageName = image.Substring(image.LastIndexOf(@"\", StringComparison.Ordinal) + 1);
                ext = Path.GetExtension(imageName).ToLower();
                if (ext == ".jpg" || ext == ".jpeg")
                {
                    switch (i)
                    {
                        case 0:
                            Image1.ImageUrl = "~/Images/" + id + "/" + imageName;
                            Image1.AlternateText = imageName;
                            break;
                        case 1:
                            Image2.ImageUrl = "~/Images/" + id + "/" + imageName;
                            Image1.AlternateText = imageName;
                            break;
                        case 2:
                            Image3.ImageUrl = "~/Images/" + id + "/" + imageName;
                            Image1.AlternateText = imageName;
                            break;
                        case 3:
                            Image4.ImageUrl = "~/Images/" + id + "/" + imageName;
                            Image1.AlternateText = imageName;
                            break;
                        default:
                            break;
                    }

                    i += 1;
                }
            }
        }
    }
}
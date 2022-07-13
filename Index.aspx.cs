using System;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Web.UI.WebControls;
using Microsoft.AspNet.Identity;

namespace ScaleModelsExcel
{
    public partial class Index : System.Web.UI.Page
    {
        private static string ConStr = string.Empty;
        private static readonly string fileName = "modelos.xls";
        private static string path = string.Empty;
        private static string ext = string.Empty;
        private static string query = string.Empty;
        private static int NumRecords = 0;
        private static bool sw_1 = false;
        private static bool sw_2 = false;
        private static bool sw_3 = false;
        readonly DataTable DtExcel = new DataTable();
        readonly DataTable DtCollection = new DataTable();
        readonly DataTable DtBrand = new DataTable();
        readonly DataTable DtMaker = new DataTable();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
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

                FillPage();
            }
            else
            {
                sw_1 = DdlCollection.SelectedIndex != 0;
                sw_2 = DdlBrand.SelectedIndex != 0;
                sw_3 = DdlMaker.SelectedIndex != 0;

                if (sw_1 || sw_2 || sw_3)
                {
                    query = "select Collection, Brand, Model, CarYear, Maker, Scale18, Scale from[Sheet1$]";

                    query += " where ";

                    if (sw_1)
                    {
                        query += "Collection = '" + DdlCollection.SelectedValue + "'";
                    }

                    if (sw_2)
                    {
                        if (sw_1)
                        {
                            query += " and ";
                        }
                        query += "Brand = '" + DdlBrand.SelectedValue + "'";
                    }

                    if (sw_3)
                    {
                        if (sw_1 || sw_2)
                        {
                            query += " and ";
                        }
                        query += "Maker = '" + DdlMaker.SelectedValue + "'";
                    }

                    query += " order by Brand, Model, CarYear, Maker, Collection;";

                }
                else
                {

                }

                string clientID = Context.User.Identity.GetUserId();

                if (clientID != null)
                {
                    LstProducts();
                }
                else
                {
                    LblResult.Text = "Please log in (or register, first) in order to see our amazing 'all timer' vehicles collection...";
                }
            }
        }

        protected void DdlCollection_SelectedIndexChanged(object sender, EventArgs e)
        {
            DdlSelChanged();
        }

        protected void DdlBrand_SelectedIndexChanged(object sender, EventArgs e)
        {
            DdlSelChanged();
        }

        protected void DdlMaker_SelectedIndexChanged(object sender, EventArgs e)
        {
            DdlSelChanged();
        }

        private void FillPage()
        {
            using (OleDbConnection con = new OleDbConnection(ConStr))
            {
                try
                {
                    con.Open();

                    query = "select Collection from[Sheet1$]";
                    OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(query, con);
                    oleDbDataAdapter.Fill(DtExcel);
                    NumRecords = DtExcel.Rows.Count;

                    query = "select distinct Collection from[Sheet1$] order by Collection;";
                    OleDbDataAdapter oleCollection = new OleDbDataAdapter(query, con);
                    oleCollection.Fill(DtCollection);

                    DdlCollection.DataSource = DtCollection;
                    DdlCollection.DataValueField = "Collection";
                    DdlCollection.DataBind();
                    DdlCollection.Items.Insert(0, new ListItem("--Select Collection--", "0"));

                    query = "select Distinct Brand from[Sheet1$] order by Brand;";
                    OleDbDataAdapter oleBrand = new OleDbDataAdapter(query, con);
                    oleBrand.Fill(DtBrand);

                    DdlBrand.DataSource = DtBrand;
                    DdlBrand.DataValueField = "Brand";
                    DdlBrand.DataBind();
                    DdlBrand.Items.Insert(0, new ListItem("--Select Brand--", "0"));

                    query = "select Distinct Maker from[Sheet1$] order by Maker;";
                    OleDbDataAdapter oleMaker = new OleDbDataAdapter(query, con);
                    oleMaker.Fill(DtMaker);

                    DdlMaker.DataSource = DtMaker;
                    DdlMaker.DataValueField = "Maker";
                    DdlMaker.DataBind();
                    DdlMaker.Items.Insert(0, new ListItem("--Select Maker--", "0"));

                    LblResult.Text = NumRecords.ToString("N0") + " models in database actually. Please, use dropdownlists for narrowing selection criteria.";
                }
                catch (Exception ex)
                {
                    LblResult.Text = "Error: " + ex;
                }
                finally
                {
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                    con.Dispose();
                }
            }

        }

        private void DdlSelChanged()
        {
            sw_1 = DdlCollection.SelectedValue == "0";
            sw_2 = DdlBrand.SelectedValue == "0";
            sw_3 = DdlMaker.SelectedValue == "0";

            string query1, query2, query3;

            if (sw_1)
            {
                if (sw_2)
                {
                    if (sw_3)
                    {
                        // case 1 - 0, 0, 0
                        query1 = "select distinct Collection from[Sheet1$] order by Collection;";
                        query2 = "select distinct Brand from[Sheet1$] order by Brand;";
                        query3 = "select distinct Maker from[Sheet1$] order by Maker;";
                    }
                    else
                    {
                        // case 4 - 0, 0, val
                        query1 = "select distinct Collection, Maker from[Sheet1$] where Maker = '" + DdlMaker.SelectedValue + "' order by Collection;";
                        query2 = "select distinct Brand, Maker from[Sheet1$] where Maker = '" + DdlMaker.SelectedValue + "' order by Brand;";
                        query3 = null;
                    }
                }
                else
                {
                    if (sw_3)
                    {
                        // case 3 - 0, val, 0
                        query1 = "select distinct Collection, Brand from[Sheet1$] where Brand = '" + DdlBrand.SelectedValue + "' order by Collection;";
                        query2 = null;
                        query3 = "select distinct Brand, Maker from[Sheet1$] where Brand = '" + DdlBrand.SelectedValue + "' order by Maker;";
                    }
                    else
                    {
                        // case 7 - 0, val, val
                        query1 = "select distinct Collection, Brand, Maker from[Sheet1$] where Brand = '" + DdlBrand.SelectedValue + "' and Maker = '" + DdlMaker.SelectedValue + "' order by Collection;";
                        query2 = null;
                        query3 = null;
                    }
                }
            }
            else
            {
                if (sw_2)
                {
                    if (sw_3)
                    {
                        // case 2 - val, 0, 0
                        query1 = null;
                        query2 = "select distinct Collection, Brand from[Sheet1$] where Collection = '" + DdlCollection.SelectedValue + "' order by Brand;";
                        query3 = "select distinct Collection, Maker from[Sheet1$] where Collection = '" + DdlCollection.SelectedValue + "' order by Maker;";
                    }
                    else
                    {
                        // case 6 - val, 0, val
                        query1 = null;
                        query2 = "select distinct Collection, Brand, Maker from[Sheet1$] where Collection = '" + DdlCollection.SelectedValue + "' and Maker = '" + DdlMaker.SelectedValue + "' order by Brand;";
                        query3 = null;
                    }
                }
                else
                {
                    if (sw_3)
                    {
                        // case 5 - val, val, 0
                        query1 = null;
                        query2 = null;
                        query3 = "select distinct Collection, Brand, Maker from[Sheet1$] where Collection = '" + DdlCollection.SelectedValue + "' and Brand = '" + DdlBrand.SelectedValue + "' order by Maker";
                    }
                    else
                    {
                        // case 8 - val, val, val
                        query1 = null;
                        query2 = null;
                        query3 = null;
                    }
                }
            }

            if (query1 != null)
            {
                Update_DdlCollection(query1);
            }

            if (query2 != null)
            {
                Update_DdlBrand(query2);
            }

            if (query3 != null)
            {
                Update_DdlMaker(query3);
            }
        }

        private void Update_DdlCollection(string q1)
        {
            using (OleDbConnection con = new OleDbConnection(ConStr))
            {
                try
                {
                    if (DtCollection.Rows.Count > 0)
                    {
                        DtCollection.Rows.Clear();
                    }

                    con.Open();
                    OleDbDataAdapter oleCollection = new OleDbDataAdapter(q1, con);
                    oleCollection.Fill(DtCollection);
                    
                    DdlCollection.DataSource = DtCollection;
                    DdlCollection.DataValueField = "Collection";
                    DdlCollection.DataBind();
                    DdlCollection.Items.Insert(0, new ListItem("--Select Collection--", "0"));

                }
                catch (Exception ex)
                {
                    LblResult.Text = "Error: " + ex;
                }
                finally
                {
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                    con.Dispose();
                }
            }
        }

        private void Update_DdlBrand(string q2)
        {
            using (OleDbConnection con = new OleDbConnection(ConStr))
            {
                try
                {
                    if (DtBrand.Rows.Count > 0)
                    {
                        DtBrand.Rows.Clear();
                    }

                    con.Open();
                    OleDbDataAdapter oleBrand = new OleDbDataAdapter(q2, con);
                    oleBrand.Fill(DtBrand);

                    DdlBrand.DataSource = DtBrand;
                    DdlBrand.DataValueField = "Brand";
                    DdlBrand.DataBind();
                    DdlBrand.Items.Insert(0, new ListItem("--Select Brand--", "0"));

                }
                catch (Exception ex)
                {
                    LblResult.Text = "Error: " + ex;
                }
                finally
                {
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                    con.Dispose();
                }
            }
        }

        private void Update_DdlMaker(string q3)
        {
            using (OleDbConnection con = new OleDbConnection(ConStr))
            {
                try
                {
                    if (DtMaker.Rows.Count > 0)
                    {
                        DtMaker.Rows.Clear();
                    }

                    con.Open();
                    OleDbDataAdapter oleMaker = new OleDbDataAdapter(q3, con);
                    oleMaker.Fill(DtMaker);

                    DdlMaker.DataSource = DtMaker;
                    DdlMaker.DataValueField = "Maker";
                    DdlMaker.DataBind();
                    DdlMaker.Items.Insert(0, new ListItem("--Select Maker--", "0"));
                    
                }
                catch (Exception ex)
                {
                    LblResult.Text = "Error: " + ex;
                }
                finally
                {
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                    con.Dispose();
                }
            }
        }
        private void LstProducts()
        {
            using (OleDbConnection con = new OleDbConnection(ConStr))
            {
                try
                {
                    if (DtExcel.Rows.Count > 0)
                    {
                        DtExcel.Rows.Clear();
                    }

                    con.Open();

                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter(query, con); //here we read data from sheet1
                    oleAdpt.Fill(DtExcel); //fill excel data into dataTable

                    if (DtExcel.Rows.Count > 0)
                    {
                        if (DtExcel.Rows.Count > 50)
                        {
                            pnlProducts.Controls.Add(new Literal { Text = "" });
                            LblResult.Text = "It's not possible to show all models, actually. Please, narrow selection criteria by using dropdownlists.";
                        }
                        else
                        {
                            LblResult.Text = "showing " + DtExcel.Rows.Count + " models matching selection criteria out of " + NumRecords.ToString("N0") + " in database actually";

                            foreach (DataRow row in DtExcel.Rows)
                            {
                                Panel productPanel = new Panel();
                                ImageButton imageButton = new ImageButton();
                                Label lblName = new Label();
                                Label lblColMaker = new Label();

                                // Set childcontrol's properties
                                if (File.Exists(Server.MapPath("~/Images/" + row["Scale18"] + "/01.jpg")))
                                {
                                    imageButton.ImageUrl = "~/Images/" + row["Scale18"] + "/01.jpg";
                                }
                                else
                                {
                                    imageButton.ImageUrl = "~/Images/nophoto.jpg";
                                };

                                imageButton.CssClass = "productImage";

                                if (!String.IsNullOrEmpty(row["Scale18"].ToString()))
                                {
                                    imageButton.PostBackUrl = "~/Pages/Product.aspx?id=" + row["Scale18"];
                                }
                                else
                                {
                                    imageButton.PostBackUrl = "";
                                }
                                imageButton.AlternateText = row["Scale18"].ToString();

                                lblName.Text = row["Brand"] + " " + row["Model"] + " (" + row["CarYear"] + ")";
                                lblName.CssClass = "productName";
                                lblColMaker.Text = row["Collection"] + " - " + row["Maker"];
                                lblColMaker.CssClass = "productName";

                                // Add Child Controls to Parent Panel
                                productPanel.Controls.Add(imageButton);
                                productPanel.Controls.Add(new Literal { Text = "<br />" });
                                productPanel.Controls.Add(lblName);
                                productPanel.Controls.Add(new Literal { Text = "<br />" });
                                productPanel.Controls.Add(lblColMaker);

                                // Add dynamic panel to static parent panle
                                pnlProducts.Controls.Add(productPanel);
                            }
                        }
                    }
                    else
                    {
                        // No products found
                        pnlProducts.Controls.Add(new Literal { Text = "no products found!" });
                    }
                }
                catch (Exception ex)
                {
                    LblResult.Text = "Error: " + ex;
                }
                finally
                {
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                    con.Dispose();
                }
            }
        }
    }
}
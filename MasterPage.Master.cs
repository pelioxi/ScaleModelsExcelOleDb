using System;
using System.Web;

namespace ScaleModelsExcel
{
    public partial class MasterPage : System.Web.UI.MasterPage
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            var user = Context.User.Identity;

            if(user.IsAuthenticated)
            {
                litStatus.Text = Context.User.Identity.Name;

                lnklogin.Visible = false;
                lnkRegister.Visible = false;

                Lnklogout.Visible = true;
                litStatus.Visible = true;
            }
            else
            {
                lnklogin.Visible = true;
                lnkRegister.Visible = true;

                Lnklogout.Visible = false;
                litStatus.Visible = false;
            }
        }

        protected void Lnklogout_Click(object sender, EventArgs e)
        {
            var authenticationManager = HttpContext.Current.GetOwinContext().Authentication;
            authenticationManager.SignOut();

            Response.Redirect("~/Index.aspx");
        }
    }
}
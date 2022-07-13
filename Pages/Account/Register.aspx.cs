using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;
using Microsoft.Owin.Security;
using ScaleModelsExcel.Model;
using System;
using System.Linq;
using System.Web;

namespace ScaleModelsExcel.Pages.Account
{
    public partial class Register : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void BtnSubmit_Click(object sender, EventArgs e)
        {
            UserStore<IdentityUser> userStore = new UserStore<IdentityUser>();

            userStore.Context.Database.Connection.ConnectionString =
                System.Configuration.ConfigurationManager.ConnectionStrings["ScaleModelsConnectionString"].ConnectionString;

            UserManager<IdentityUser> manager = new UserManager<IdentityUser>(userStore);

            IdentityUser user = new IdentityUser
            {
                UserName = TxtUserName.Text
            };

            if (TxtPassword.Text != TxtConfirmPassword.Text)
            {
                LitStatus.Text = "Passwords must match";
            }
            else
            {
                try
                {
                    IdentityResult result = manager.Create(user, TxtPassword.Text);

                    if (result.Succeeded)
                    {

                        UserInformation info = new UserInformation
                        {
                            Address = TxtAddress.Text,
                            FirstName = TxtFirstName.Text,
                            LastName = TxtLastName.Text,
                            PostalCode = Convert.ToInt32(TxtPostalCode.Text),
                            GUID = user.Id
                        };

                        UserInfoModel model = new UserInfoModel();
                        model.InsertUserInformation(info);

                        var authenticationManager = HttpContext.Current.GetOwinContext().Authentication;

                        var userIdentity = manager.CreateIdentity(user, DefaultAuthenticationTypes.ApplicationCookie);

                        authenticationManager.SignIn(new AuthenticationProperties(), userIdentity);
                        Response.Redirect("~/Index.aspx");
                    }
                    else
                    {
                        LitStatus.Text = result.Errors.FirstOrDefault();
                    }
                }
                catch (Exception ex)
                {
                    LitStatus.Text = ex.ToString();
                }
            }
        }
    }
}
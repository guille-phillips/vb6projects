using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;


public partial class Controls_MainMenu : System.Web.UI.UserControl
{
    private const string SELECTED_CSS = "Selected";

    protected void Page_Load(object sender, EventArgs e)
    {
        // Determine which menu item must be selected by looking at the filename for the current request.
        switch (Request.AppRelativeCurrentExecutionFilePath.ToLower())
        {
            case "~/default.aspx":
                lnkHome.CssClass = SELECTED_CSS; break;

            case "~/checkavailability.aspx":
                lnkCheckAvailability.CssClass = SELECTED_CSS; break;

            case "~/createappointment.aspx":
                lnkMakeAppointment.CssClass = SELECTED_CSS; break;

            case "~/signup.aspx":
                lnkSignUp.CssClass = SELECTED_CSS; break;

            case "~/login.aspx":
                if (Context.User.Identity.IsAuthenticated == false)
                {
                  // Only do this when the user is not logged in. Otherwise, the Logout link is
                  // visible, which cannot be given the Selected style, because the user is redirected to the
                  // home page after logging out.
                    HyperLink myHyperLink =(HyperLink) Convert.ChangeType(lvLogin.FindControl("lnkLogin"), typeof(HyperLink));
                    myHyperLink.CssClass = SELECTED_CSS;
                }
                break;

            default:
                
                // Handle all other pages in the Management section
                if (Request.AppRelativeCurrentExecutionFilePath.ToLower().Contains("~/management"))
                {
                    lnkManagement.CssClass = SELECTED_CSS;
                }
                break;

        }
      
    }
}

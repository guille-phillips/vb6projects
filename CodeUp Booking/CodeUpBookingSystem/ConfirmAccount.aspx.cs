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

public partial class ConfirmAccount : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            // Get the user's Guid from the Query String.
            Guid userId = new Guid(Request.QueryString.Get("Id"));

            MembershipUser myUser = Membership.GetUser(userId);

            if (myUser != null)
            {
                //The user was found in the system, so we can activate the account.
                myUser.IsApproved = true;
                Membership.UpdateUser(myUser);
                plcSuccess.Visible = true;
            }
            else
            {
                lblErrorMessage.Visible = true;
            }

        }
        catch (Exception ex)
        { lblErrorMessage.Visible = true; }
    }
}

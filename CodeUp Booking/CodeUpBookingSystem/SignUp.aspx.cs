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

public partial class SignUp : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    
    protected void CreateUserWizard1_ContinueButtonClick(object sender, EventArgs e)
    {
        Response.Redirect("~/");
    }
    protected void CreateUserWizard1_SendingMail(object sender, MailMessageEventArgs e)
    {
        //Customize the mail body by replacing the placeholders in the static mail body file with actual values.
        e.Message.Body = e.Message.Body.Replace("##Id##", Membership.GetUser(CreateUserWizard1.UserName).ProviderUserKey.ToString());
        e.Message.Body = e.Message.Body.Replace("##UserName##", CreateUserWizard1.UserName);

        string applicationFolder = Request.ServerVariables.Get("SCRIPT_NAME").Substring(0, Request.ServerVariables.Get("SCRIPT_NAME").LastIndexOf("/"));
        e.Message.Body = e.Message.Body.Replace("##FullRootUrl##", Helpers.GetCurrentServerRoot + applicationFolder);
      
    }
}

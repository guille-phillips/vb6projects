using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

/// <summary>
/// The Helpers class contains helper members that can be used throughout the application without being tied to the Business or Data Access layer.
/// </summary>
public class Helpers
{
	public Helpers()
	{
		//
		// TODO: Add constructor logic here
		//
	}
    
  // <summary>
  // Returns a string with the current protocol and server name. E.g. http://www.SomeSite.nl or http://localhost:2175
  // </summary>
  // <remarks>
  // This property is useful if you want to know the current site in a dynamic environment, 
  // as is the case with the ASP.NET Development Server with dynamic ports or with
  // IIS in a site with multiple host header definitions.
  //</remarks>
  public static string GetCurrentServerRoot
  {
    get
      {
      string protocol =string.Empty;
      string siteName;

      if (HttpContext.Current.Request.ServerVariables.Get("SERVER_PORT") != "1")
      {protocol = "http://";}
      else
      {protocol = "https://";}
      
      siteName = HttpContext.Current.Request.ServerVariables.Get("HTTP_HOST");

      return protocol + siteName;
      }
  }
}

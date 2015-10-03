using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Data.SqlClient;


/// <summary>
/// Summary description for AppConfiguration
/// </summary>
public class AppConfiguration
{
    private static string _applicationIdString = string.Empty;
	
    //Returns the connection string for the Booking system.
    public static string ConnectionString
    {
        get
        {
            string tempValue = @"server=(local)\SqlExpress;AttachDbFileName=|DataDirectory|aspnetdb.mdf;Integrated Security=true;User Instance=true";
            try
            {
                if (ConfigurationManager.ConnectionStrings["LocalSqlServer"] != null)
                    tempValue = ConfigurationManager.ConnectionStrings["LocalSqlServer"].ConnectionString;
            }
            catch
            {
                /* When we can't get the settting from the config file, ignore the error and return 
                    the default connection string that points to the instance SqlExpress on the local 
                machine and tries to attach the database AppointmentBooking automatically.*/
            }
            return tempValue;
        }
    }

    //Returns the application id for the Booking system.
    public static string ApplicationID
    {
        get
        {
            //we store _applicationIdString as a private static so we only ever access the b once for the AppId
            if (_applicationIdString == string.Empty)
            {
                using (SqlConnection conn = new SqlConnection(AppConfiguration.ConnectionString))
                {
                    conn.Open();
                    const String selectQuery = "aspnet_Application_FindApplicationByName";
                    using (SqlCommand cmd = new SqlCommand(selectQuery, conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        SqlParameter p1 = cmd.Parameters.Add("@ApplicationName", System.Data.SqlDbType.NVarChar);
                        SqlParameter p2 = cmd.Parameters.Add("@ApplicationId", System.Data.SqlDbType.UniqueIdentifier);
                        p2.Direction = ParameterDirection.Output;
                        p1.Value = Membership.ApplicationName;
                        p2.Value = null;

                        cmd.ExecuteNonQuery();

                        _applicationIdString = cmd.Parameters["@ApplicationId"].Value.ToString();
                    }
                    if (conn != null)
                        conn.Close();
                }
            }
            return _applicationIdString;
        }
    }

    
  // <summary>
  // gets the value that determines whether comments are required in an appointment request.
  // </summary>
  public static bool RequireCommentsInRequest
  {
    get
    {
      bool tempValue = true;
      try
      {
        if (ConfigurationManager.AppSettings.Get("RequireCommentsInRequest")!=null)
          tempValue = Convert.ToBoolean(ConfigurationManager.AppSettings.Get("RequireCommentsInRequest"));
      }
      catch{}
        //Ignore the error, and return the default value for tempValue
      
      return tempValue;

    }
  }

  // <summary>
  // gets the singular form of the user-friendly name of the Booking Object.
  // </summary>
  public static string BookingObjectNameSingular
  {
    get
    {
       string tempValue = string.Empty;
      try
      {
        if (ConfigurationManager.AppSettings.Get ("BookingObjectNameSingular")!=null)
          tempValue = ConfigurationManager.AppSettings.Get("BookingObjectNameSingular");
      }
      catch{
        //Pass up the error
        throw new Exception("Name not set");
        }
      return tempValue;
      }
  }

  // <summary>
  // gets the plural form of the user-friendly name of the Booking Object.
  // </summary>
  public static string BookingObjectNamePlural
  {
    get
    {
       string tempValue = string.Empty;
      try
      {
          if (ConfigurationManager.AppSettings.Get("BookingObjectNamePlural") != null)
              tempValue = ConfigurationManager.AppSettings.Get("BookingObjectNamePlural");
      }
      catch{
        //Pass up the erorr
        throw new Exception("Name not set");
      }
      return tempValue;
      }
  }

}

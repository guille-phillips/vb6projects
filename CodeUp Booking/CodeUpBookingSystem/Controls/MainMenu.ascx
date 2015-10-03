<%@ Control Language="C#" AutoEventWireup="true" CodeFile="MainMenu.ascx.cs" Inherits="Controls_MainMenu" %>
<ul>
  <li>
    <asp:HyperLink ID="lnkHome" runat="server" NavigateUrl="~/Default.aspx">Home</asp:HyperLink></li>
  <li>
    <asp:HyperLink ID="lnkCheckAvailability" runat="server" NavigateUrl="~/CheckAvailability.aspx">Check Availability</asp:HyperLink></li>
  <li>
    <asp:HyperLink ID="lnkMakeAppointment" runat="server" NavigateUrl="~/CreateAppointment.aspx">Make Appointment</asp:HyperLink></li>
  <li>
    <asp:HyperLink ID="lnkSignUp" runat="server" NavigateUrl="~/SignUp.aspx">Sign Up</asp:HyperLink></li>
  <asp:LoginView runat="server" ID="lvLogin">
    <AnonymousTemplate>
      <li>
        <asp:HyperLink ID="lnkLogin" runat="server" NavigateUrl="~/Login.aspx">Login</asp:HyperLink></li>
    </AnonymousTemplate>
    <LoggedInTemplate>
      <li>
        <asp:HyperLink ID="lnkLogout" runat="server" NavigateUrl="~/Logout.aspx">Logout</asp:HyperLink></li>
    </LoggedInTemplate>
  </asp:LoginView>
  <li>
    <asp:HyperLink ID="lnkManagement" runat="server" NavigateUrl="~/Management/Default.aspx">Management</asp:HyperLink></li>
</ul>

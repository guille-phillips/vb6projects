<%@ Control Language="C#" AutoEventWireup="true" CodeFile="ManagementMenu.ascx.cs" Inherits="Controls_ManagementMenu" %>
<ul>
  <li>
    <asp:HyperLink ID="lnkAppointments" runat="server" NavigateUrl="~/Management/Appointments.aspx">Appointments</asp:HyperLink>
  </li>
  <li>
    <asp:HyperLink ID="lnkBookingObjects" runat="server" Text="<%$ AppSettings:BookingObjectNamePlural %>" NavigateUrl="~/Management/BookingObjects.aspx"></asp:HyperLink>
  </li>
  <li>
    <asp:HyperLink ID="lnkConfiguration" runat="server" NavigateUrl="~/Management/Configuration.aspx">Application Settings </asp:HyperLink>
  </li>
</ul>
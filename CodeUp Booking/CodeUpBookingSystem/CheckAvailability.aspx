<%@ Page Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="CheckAvailability.aspx.cs" Inherits="CheckAvailability" Title="Untitled Page" %>
<%@ Register Src="Controls/TimeSheet.ascx" TagName="TimeSheet" TagPrefix="CodeUp" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
  <h1>Availability Checker</h1>
  <p>
    You can use the Availability Checker to see if your favorite
    <asp:Literal ID="Literal1" runat="server" Text="<%$ AppSettings:BookingObjectNameSingular %>"></asp:Literal>
    is available at the date of your choice. Click the Calendar icon below and select the date you want for your appointment. You'll see a list with all the
    <asp:Literal ID="Literal2" runat="server" Text="<%$ AppSettings:BookingObjectNamePlural %>"></asp:Literal>
    and an indication of their availability. When the
    <asp:Literal ID="Literal3" runat="server" Text="<%$ AppSettings:BookingObjectNameSingular %>"></asp:Literal>
    is available, you can make an appointment directly by clicking &quot;Book&quot;.
  </p>
  <asp:Label ID="lblSelectedDate" runat="server" Text="Please select a date:" Visible="false" /><asp:Label ID="lblInstructions" runat="server" Text="Please select a date:" />
  <a href="#">
    <img src="Images/Calendar.gif" onclick="ToggleDisplay('ctl00_ContentPlaceHolder1_divCalendar');" align="middle" /></a>
  <br />
  <div id="divCalendar" runat="server">
    <asp:Calendar ID="calAppointmentDate" runat="server"></asp:Calendar>
  </div>
  <asp:CustomValidator ID="valSelectedDate" runat="server" CssClass="ErrorMessage" ErrorMessage="You cannot check the availability for past dates." ForeColor=""></asp:CustomValidator>
  <br />
  <CodeUp:TimeSheet ID="TimeSheet1" runat="server" StartTime="<%$ AppSettings:FirstAvailableWorkingHour %>" EndTime="<%$ AppSettings:LastAvailableWorkingHour %>" />
</asp:Content>



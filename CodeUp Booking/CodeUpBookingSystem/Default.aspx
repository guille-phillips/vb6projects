<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/MasterPage.master"  CodeFile="Default.aspx.cs" Inherits="_Default" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <h1>
    Welcome to the ... Booking System</h1>
  <p>
    This application allows you to quickly book your favourite
    <asp:Literal ID="Literal1" runat="server" Text="<%$ AppSettings:BookingObjectNameSingular %>"></asp:Literal>.<br />
    Before you can use this application, you need to have an active account. If you don't have an account yet, you can <a href="SignUp.aspx">create one now</a>. Otherwise, you'll be asked to login before you can make the appointment.
  </p>
  <p>
    Proceed to <a href="CreateAppointment.aspx">make an appointment</a>, or look at the <a href="CheckAvailability.aspx">Availability Checker</a> to see if your favourite
    <asp:Literal ID="Literal2" runat="server" Text="<%$ AppSettings:BookingObjectNameSingular %>"></asp:Literal>
    is available.</p>
</asp:Content>

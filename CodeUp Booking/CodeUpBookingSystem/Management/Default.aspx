<%@ Page Language="C#" MasterPageFile="~/ManagementMaster.master" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="Management_Default" Title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
  <h1>
    Appointment Booking System - Management Section</h1>
  <p>
    This is the Management section for the Appointment Booking System that allows you to view a list of all the <a href="Appointments.aspx">available appointments</a>, manage the <a href="BookingObjects.aspx">
      <asp:Literal ID="Literal1" runat="server" Text="<%$ AppSettings:BookingObjectNamePlural %>"></asp:Literal></a> (create new and change existing objects, and change their availability) and to manage the <a href="Configuration.aspx">application settings</a> like the user-friendly name of the Booking Objects.
  </p>
</asp:Content>


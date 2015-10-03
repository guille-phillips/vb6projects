<%@ Page Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="CreateAppointment.aspx.cs" Inherits="CreateAppointment" Title="Untitled Page" %>

<%@ Register Src="~/Controls/TimePicker.ascx" TagName="TimePicker" TagPrefix="CodeUp" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <asp:Wizard ID="wizAppointment" runat="server" Width="700px" ActiveStepIndex="0">
    <HeaderStyle Width="300px" />
    <StepStyle VerticalAlign="Top" />
    <SideBarStyle VerticalAlign="Top" Width="200px" />
    <WizardSteps>
      <asp:WizardStep ID="WizardStep1" runat="server" StepType="Start" Title="Introduction">
        <h1>
          <asp:Literal ID="Literal1" runat="server" Text="<%$ AppSettings:BookingObjectNameSingular %>"></asp:Literal>
          Selection Wizard</h1>
        Welcome to the
        <asp:Literal ID="Literal2" runat="server" Text="<%$ AppSettings:BookingObjectNameSingular %>"></asp:Literal>
        Selection Wizard. This wizard will guide you through the process of making an appointment. Click Next to continue.</asp:WizardStep>
      <asp:WizardStep ID="WizardStep2" runat="server" Title="Select" StepType="Step">
        <h1>
          Select
          <asp:Literal ID="Literal5" runat="server" Text="<%$ AppSettings:BookingObjectNameSingular %>"></asp:Literal></h1>
        Select a
        <asp:Literal ID="Literal3" runat="server" Text="<%$ AppSettings:BookingObjectNameSingular %>"></asp:Literal>
        for your appointment from the drop-down list and click Next.<br />
        <br />
        <asp:DropDownList ID="lstBookingObject" runat="server" DataSourceID="odsBookingObjectList" DataTextField="Title" DataValueField="Id">
        </asp:DropDownList>
        <br />
        <asp:RequiredFieldValidator ID="reqBookingObject" runat="server" ControlToValidate="lstBookingObject"></asp:RequiredFieldValidator>
        <br />
        <asp:ObjectDataSource ID="odsBookingObjectList" runat="server" SelectMethod="GetBookingObjectList" TypeName="BookingObjectManager"></asp:ObjectDataSource>
      </asp:WizardStep>
      <asp:WizardStep ID="WizardStep3" runat="server" Title="Select Date" StepType="Step">
        <h1>
          Select Date
        </h1>
        Select a date for your appointment from the calendar and click Next to continue.<br />
        <br />
        <asp:Calendar ID="calStartDate" runat="server" />
        <br />
        <asp:Literal ID="litSelectedDate" runat="server"></asp:Literal>
        <asp:CustomValidator ID="valStartDate1" runat="server" Display="Dynamic" ErrorMessage="Please select a date" ForeColor="" CssClass="ErrorMessage"></asp:CustomValidator>
        <asp:CustomValidator ID="valStartDate2" runat="server" Display="Dynamic" ErrorMessage="Your appointment date cannot lay in the past." ForeColor="" CssClass="ErrorMessage"></asp:CustomValidator>
        <br />
        <br />
      </asp:WizardStep>
      <asp:WizardStep ID="WizardStep4" runat="server" StepType="Step" Title="Select Time">
        <h1>
          Select Time</h1>
        <br />
        Select a start time and a duration for your appointment and click Next to continue.<br />
        <br />
        Start time
        <CodeUp:TimePicker ID="hpTime" runat="server" StartTime="<%$ AppSettings:FirstAvailableWorkingHour %>" EndTime="<%$ AppSettings:LastAvailableWorkingHour %>" />
        &nbsp; Duration
        <asp:DropDownList ID="lstDuration" runat="server">
          <asp:ListItem>1</asp:ListItem>
          <asp:ListItem>2</asp:ListItem>
          <asp:ListItem>3</asp:ListItem>
          <asp:ListItem>4</asp:ListItem>
          <asp:ListItem>5</asp:ListItem>
          <asp:ListItem>6</asp:ListItem>
          <asp:ListItem>7</asp:ListItem>
          <asp:ListItem>8</asp:ListItem>
        </asp:DropDownList>
        &nbsp;hour<br />
        <br />
      </asp:WizardStep>
      <asp:WizardStep ID="WizardStep5" runat="server" Title="Comments" StepType="Step">
        <h1>
          Enter Comments</h1>
        Enter any comments you want to add to your appointment request and click Next to continue.<br />
        <br />
        <asp:TextBox ID="txtComments" runat="server" Height="179px" TextMode="MultiLine" Width="430px"></asp:TextBox>
        <br />
        <asp:RequiredFieldValidator ID="reqComments" runat="server" ControlToValidate="txtComments" ErrorMessage="Please enter a comment"></asp:RequiredFieldValidator>
      </asp:WizardStep>
      <asp:WizardStep ID="WizardStep6" runat="server" StepType="Finish" Title="Review Your Request">
        <h1>
          Review your Request</h1>
        Please review the options you selected below. If you're sure all details are filled in correctly, click Finish to finalize your appointment request. Otherwise, click the Previous button to make any changes.<br /><br />
        <table>
        <tr>
          <td style="width: 150px" class="Label">
            <asp:Literal ID="Literal4" runat="server" Text="<%$ AppSettings:BookingObjectNameSingular %>"></asp:Literal>:</td>
          <td>
            <asp:Label ID="lblBookingObject" runat="server" Text="Label"></asp:Label>
          </td>
        </tr>
        <tr>
          <td class="Label">
            Date and time</td>
          <td>
            From
            <asp:Label ID="lblStartTime" runat="server" Text="Label"></asp:Label>
            till
            <asp:Label ID="lblEndTime" runat="server" Text="Label"></asp:Label>
          </td>
        </tr>
        <tr>
          <td class="Label">
            Comments:</td>
          <td>
            <asp:Literal ID="lblComments" runat="server"></asp:Literal>
          </td>
        </tr>
        </table>
      </asp:WizardStep>
    </WizardSteps>
  </asp:Wizard>
  <br />
  <asp:MultiView ID="MultiView1" runat="server">
    <asp:View ID="ViewSuccess" runat="server">
      Your appointment has been made.</asp:View>
    <asp:View ID="ViewFailure" runat="server">
      Sorry, the date and time you selected are not available. You can
      <asp:LinkButton ID="lnkTryAgain" runat="server">try again</asp:LinkButton>, or find an available time with the <a href="CheckAvailability.aspx">availability checker</a>.</asp:View>
  </asp:MultiView>
</asp:Content>

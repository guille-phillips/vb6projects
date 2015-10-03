<%@ Page Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="SignUp.aspx.cs" Inherits="SignUp" Title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
  <h1>
    Sign up for an Account with the ... Booking System</h1>
  <p>
    Before you can use the ... Booking System to make an appointment, you'll need to create an account first and then confirm your email address.</p>
  <p>
    Enter your details below and click the Sign Up button. You'll receive an email with instructions about confirming your email address and activating your account.</p>
  <asp:CreateUserWizard ID="CreateUserWizard1" runat="server" 
        DisableCreatedUser="True" CreateUserButtonText="Sign Up" 
        LoginCreatedUser="False" 
        CompleteSuccessText="Your account has been successfully created. You'll receive an e-mail with instructions about activating your account shortly." 
        oncontinuebuttonclick="CreateUserWizard1_ContinueButtonClick" 
        onsendingmail="CreateUserWizard1_SendingMail">
    <WizardSteps>
      <asp:CreateUserWizardStep ID="CreateUserWizardStep1" runat="server" Title="">
      </asp:CreateUserWizardStep>
      <asp:CompleteWizardStep ID="CompleteWizardStep1" runat="server" Title="">
      </asp:CompleteWizardStep>
    </WizardSteps>
    <MailDefinition BodyFileName="~/StaticFiles/OptInEmail.html" From="Appointment Booking &lt;You@YourProvider.Com&gt;" IsBodyHtml="True" Subject="Please Confirm Your Account With the Appointment Booking System">
    </MailDefinition>
  </asp:CreateUserWizard>



</asp:Content>


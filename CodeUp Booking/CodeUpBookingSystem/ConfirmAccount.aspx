<%@ Page Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="ConfirmAccount.aspx.cs" Inherits="ConfirmAccount" Title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
  <asp:Label ID="lblErrorMessage" runat="Server" CssClass="ErrorMessage" Text="An error occurred while trying to activate your account" Visible="False" />
  <asp:PlaceHolder ID="plcSuccess" runat="Server" Visible="False">Your account has been confirmed successfully. You can now proceed to the <a href="Login.aspx">Login</a> page and start making appointments.</asp:PlaceHolder>
</asp:Content>


<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.Master" AutoEventWireup="true" CodeBehind="Register.aspx.cs" Inherits="ScaleModelsExcel.Pages.Account.Register" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:Literal ID="LitStatus" runat="server"></asp:Literal>
    <p>
        User Name:</p>
    <asp:TextBox ID="TxtUserName" runat="server" CssClass="inputs"></asp:TextBox>
    <p>
        Password:</p>
    <asp:TextBox ID="TxtPassword" runat="server" TextMode="Password" CssClass="inputs"></asp:TextBox>
    <p>
        Confirm Password:</p>
    <asp:TextBox ID="TxtConfirmPassword" runat="server" TextMode="Password" CssClass="inputs"></asp:TextBox>

    <p>First Name:</p>
    <asp:TextBox ID="TxtFirstName" runat="server" CssClass="inputs"></asp:TextBox>

    <p>Last Name:</p>
        <asp:TextBox ID="TxtLastName" runat="server" CssClass="inputs"></asp:TextBox>

    <p>Address:</p>
        <asp:TextBox ID="TxtAddress" runat="server" CssClass="inputs"></asp:TextBox>

    <p>Zip Code:</p>
        <asp:TextBox ID="TxtPostalCode" runat="server" CssClass="inputs"></asp:TextBox>
    <br />
    <asp:Button ID="BtnSubmit" runat="server" Text="Button" CssClass="button" OnClick="BtnSubmit_Click" />
</asp:Content>

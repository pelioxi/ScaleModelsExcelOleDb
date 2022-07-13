<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.Master" AutoEventWireup="true" CodeBehind="Login.aspx.cs" Inherits="ScaleModelsExcel.Pages.Account.Login" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:Literal ID="LitStatus" runat="server"></asp:Literal>
    <p>
        User Name:</p>
    <p>
    <asp:TextBox ID="TxtUserName" runat="server" CssClass="inputs"></asp:TextBox>
    </p>
    <p>
        Password:</p>
    <p>
    <asp:TextBox ID="TxtPassword" runat="server" TextMode="Password" CssClass="inputs"></asp:TextBox>
    </p>
    <p>
    <asp:Button ID="BtnSubmit" runat="server" Text="Button" CssClass="button" OnClick="BtnSubmit_Click" />
    </p>
</asp:Content>

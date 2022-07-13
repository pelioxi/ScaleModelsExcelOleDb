<%@ Page Title="Scale Models" Language="C#" MasterPageFile="~/MasterPage.Master" AutoEventWireup="true" CodeBehind="Index.aspx.cs" Inherits="ScaleModelsExcel.Index" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div align="center">
        <table id="TblDdl" cellspacing="4">
            <tr>
                <td>Collection: </td>
                <td>
                    <asp:DropDownList ID="DdlCollection" runat="server" OnSelectedIndexChanged="DdlCollection_SelectedIndexChanged" AutoPostBack="true"/>
                </td>
                <td>Brand: </td>
                <td>
                    <asp:DropDownList ID="DdlBrand" runat="server" OnSelectedIndexChanged="DdlBrand_SelectedIndexChanged" AutoPostBack="true" />
                </td>
                <td>Maker: </td>
                <td>
                    <asp:DropDownList ID="DdlMaker" runat="server" OnSelectedIndexChanged="DdlMaker_SelectedIndexChanged" AutoPostBack="true" />
                </td>
            </tr>
        </table>
        <br />
        <asp:Label ID="LblResult" runat="server" Text="" Font-Bold="True" />
        <br />
        <br />
        <table>
            <tr>
                <td>
                    <asp:Panel ID="pnlProducts" runat="server" />
                </td>
            </tr>
        </table>
    </div>
    <div style="clear:both"></div>
    <br />
    
</asp:Content>

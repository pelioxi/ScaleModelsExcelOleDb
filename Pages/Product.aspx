<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.Master" AutoEventWireup="true" CodeBehind="Product.aspx.cs" Inherits="ScaleModelsExcel.Pages.Product" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .auto-style1 {
            height: 23px;
        }
        .auto-style2 {
            width: 107px;
        }
        .auto-style3 {
            height: 23px;
            width: 107px;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <br />
    <asp:Literal ID="LitStatus" runat="server" Text=""></asp:Literal>
    <br />
    <br />
    <table width="40%" id="detailsPage" style="vertical-align: top">
        <tr>
            <td class="auto-style2" style="font-weight: bold">
                Collection:
            </td>
            <td>
                <asp:Label ID="Label1" runat="server" Text=""></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="auto-style3" style="font-weight: bold">Brand:</td>
            <td class="auto-style1">
                <asp:Label ID="Label2" runat="server" Text=""></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="auto-style2" style="font-weight: bold">Model:</td>
            <td>
                <asp:Label ID="Label3" runat="server" Text=""></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="auto-style2" style="font-weight: bold">Car Year:</td>
            <td>
                <asp:Label ID="Label4" runat="server" Text=""></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="auto-style3" style="font-weight: bold">Maker:</td>
            <td class="auto-style1">
                <asp:Label ID="Label5" runat="server" Text=""></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="auto-style2" style="font-weight: bold">Scale:</td>
            <td>
                <asp:Label ID="Label6" runat="server" Text=""></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="auto-style2" style="font-weight: bold">Part Number:</td>
            <td>
                <asp:Label ID="Label7" runat="server" Text=""></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="auto-style2" style="font-weight: bold">Car Number:</td>
            <td>
                <asp:Label ID="Label8" runat="server" Text=""></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="auto-style2" style="font-weight: bold">Colour/Sponsor:</td>
            <td>
                <asp:Label ID="Label9" runat="server" Text=""></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="auto-style2" style="font-weight: bold">Driver:</td>
            <td>
                <asp:Label ID="Label10" runat="server" Text=""></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="auto-style2" style="font-weight: bold">Details:</td>
            <td>
                <asp:TextBox ID="TextBox1" runat="server" TextMode="MultiLine"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td class="auto-style2" style="font-weight: bold">Model Date:</td>
            <td>
                <asp:Label ID="Label11" runat="server" Text=""></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="auto-style2" style="font-weight: bold">Serial:</td>
            <td>
                <asp:Label ID="Label12" runat="server" Text=""></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="auto-style2" style="font-weight: bold">Limited Edition:</td>
            <td>
                <asp:Label ID="Label13" runat="server" Text=""></asp:Label>
            </td>
        </tr>
            <tr>
            <td class="auto-style2" style="font-weight: bold">Comments:</td>
            <td>
                <asp:TextBox ID="TextBox2" runat="server" TextMode="MultiLine"></asp:TextBox>
            </td>
        </tr>
    </table>    
    <br />
    <br />
    <table width="100%">
        <tr>
            <td rowspan="2"></td>
        </tr>
        <tr style="width:80%">
            <td id="TdImages">
                <asp:Image ID="Image1" runat="server" Height="600px" Width="100%" /><br />
                <asp:Image ID="Image2" runat="server" Height="600px" Width="100%" /><br />
                <asp:Image ID="Image3" runat="server" Height="600px" Width="100%" /><br />
                <asp:Image ID="Image4" runat="server" Height="600px" Width="100%" />
            </td>
        </tr>
    </table>
</asp:Content>

﻿<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Excel to SQL Server</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Button ID="BtnUploadFile" runat="server" Text="Upload File" OnClick="BtnUploadFile_Click"/>
        <asp:Label ID="Label1" runat="server" Text="" ForeColor="Green"></asp:Label>
    </div>
    </form>
</body>
</html>

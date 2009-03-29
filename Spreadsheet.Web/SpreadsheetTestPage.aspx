<%@ Page Language="C#" AutoEventWireup="true" %>

<%@ Register Assembly="System.Web.Silverlight" Namespace="System.Web.UI.SilverlightControls"
    TagPrefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Spreadsheet</title>
    <style type="text/css">
    body {
      padding: 0;
      margin: 0;
    }
    </style>
    <link rel="Stylesheet" href="Content/Site.css" />
</head>
<body>
    <form id="form1" runat="server" style="height:500">
        <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>
        <div>
            <asp:Silverlight ID="Silverlight1" runat="server" Source="~/ClientBin/Spreadsheet.xap" Windowless="true" MinimumVersion="3.0.40307.0" Width="100%" Height="360" />
        </div>
    </form>
  <style type="text/css">
    div.silverlightDlrWindow,
    div.silverlightDlrWindowMenu,
    #silverlightDlrRepl, 
    #silverlightDlrReplCode, 
    #silverlightDlrReplResult, 
    .silverlightDlrReplPrompt {
      font-size: 20px;
    }
    
    div.silverlightDlrWindow {
      bottom: 40px;
      border: 4px solid black;
    }

    div.silverlightDlrWindowMenu {
      height: 30px;
    }

    div.silverlightDlrWindowMenu a.active {
      border-bottom: 4px solid black;
      border-left: 4px solid black;
      border-right: 4px solid black;
      padding-top: 9px;
    }

    #silverlightDlrRepl {
      padding: 10px;
    }

    #silverlightDlrRepl, 
    #silverlightDlrReplCode, 
    #silverlightDlrReplResult, 
    .silverlightDlrReplPrompt {
      line-height: 24px;
    }

    #silverlightDlrReplCode {
      height: 24px;
    }
  </style>
</body>
</html>

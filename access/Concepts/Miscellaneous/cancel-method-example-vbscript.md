---
title: Cancel method example (VBScript)
ROBOTS: INDEX
ms.prod: access
ms.assetid: 3c5a14fa-f4b1-6c32-9014-505817c6e4cf
ms.date: 06/08/2017
localization_priority: Normal
---


# Cancel method example (VBScript)

**Applies to:** Access 2013 | Access 2016

The following example shows how to read the [Cancel](https://msdn.microsoft.com/library/747edc04-a5cc-3631-2d0b-82e7e41a76b7%28Office.15%29.aspx) method at run time. Cut and paste the following code to Notepad or another text editor and save it as **CancelVBS.asp**. You can view the result in any client browser.

```vb

<!-- BeginCancelVBS --><Script Language="VBScript">
<!--Sub cmdCancelAsync_OnClick
' Terminates currently running AsyncExecute,' ReadyState property set to adcReadyStateLoaded,
' Recordset set to NothingADC.Cancel
End Sub 
Sub cmdRefreshTable_OnClickADC.Refresh
End Sub-->
</Script> 
<OBJECT CLASSID="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33" ID="ADC">.
<PARAM NAME="SQL" VALUE="Select FirstName, LastName from Employees"><PARAM NAME="CONNECT" VALUE="Provider='sqloledb';Integrated Security='SSPI';Initial Catalog='Northwind'">
<PARAM NAME="Server" VALUE="https://<%=Request.ServerVariables("SERVER_NAME")%>">.
</OBJECT> 
<TABLE DATASRC=#ADC><TBODY>
<TR><TD><SPAN DATAFLD="FirstName"></SPAN></TD>
<TD><SPAN DATAFLD="LastName"></SPAN></TD></TR>
</TBODY></TABLE> 
<FORM><INPUT type="button" value="Refresh" id=cmdRefreshTable name=cmdRefreshTable>
<INPUT type="button" value="Cancel" id=cmdCancelAsync name=cmdCancelAsync></FORM>
<!-- EndCancelVBS -->
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
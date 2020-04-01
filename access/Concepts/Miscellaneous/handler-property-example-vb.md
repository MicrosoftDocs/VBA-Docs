---
title: Handler property example (VB)
ROBOTS: INDEX
ms.prod: access
ms.assetid: e401e7b2-754b-a66c-bfcc-8f6e3966a908
ms.date: 06/08/2019
localization_priority: Normal
---


# Handler property example (VB)

**Applies to:** Access 2013 | Access 2016

This example demonstrates the [RDS DataControl](https://msdn.microsoft.com/library/ac430669-7628-696c-c036-b5d35405d788%28Office.15%29.aspx) object [Handler](https://msdn.microsoft.com/library/aaf8c8c6-f95b-3cf3-b3f6-203f37464c87%28Office.15%29.aspx) property. (See [DataFactory Customization](https://msdn.microsoft.com/library/43cd7416-1f05-87ee-22f0-6cf0d2d1b39f%28Office.15%29.aspx) for more details.)

Assume that the following sections in the parameter file, Msdfmap.ini, are located on the server:

```ini
[connect AuthorDataBase] 
Access=ReadWrite 
Connect="DSN=Pubs" 
[sql AuthorById] 
SQL="SELECT * FROM Authors WHERE au_id = ?" 

```

Your code looks like the following. The command assigned to the [SQL](sql-property-ado.md) property will match the **_AuthorById_** identifier and will retrieve a row for author Michael O'Leary. The **DataControl** object **Recordset** property is assigned to a disconnected [Recordset](https://msdn.microsoft.com/library/0f963bf8-f066-dc8a-b754-f427de712df1%28Office.15%29.aspx) object purely as a coding convenience.

```vb
'BeginHandlerVBPublic Sub Main()
On Error GoTo ErrorHandlerDim dc As New DataControl
Dim rst As ADODB.Recordsetdc.Handler = "MSDFMAP.Handler"
dc.ExecuteOptions = 1dc.FetchOptions = 1
dc.Server = "https://MyServer"dc.Connect = "Data Source=AuthorDataBase"
dc.SQL = "AuthorById('267-41-2394')"dc.Refresh 'Retrieve the record
Set rst = dc.Recordset 'Use another Recordset as a convenienceDebug.Print "Author is '" & rst!au_fname & " " & rst!au_lname & "'"
' clean upIf rst.State = adStateOpen Then rst.Close
Set rst = NothingSet dc = Nothing
Exit SubErrorHandler:
' clean upIf Not rst Is Nothing Then
If rst.State = adStateOpen Then rst.CloseEnd If
Set rst = NothingSet dc = Nothing
If Err <> 0 ThenMsgBox Err.Source & "-->" & Err.Description, , "Error"
End IfEnd Sub
'EndHandlerVB
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
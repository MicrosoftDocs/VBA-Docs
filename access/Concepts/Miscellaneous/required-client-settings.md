---
title: Required client settings
ROBOTS: INDEX
ms.prod: access
ms.assetid: edd196b2-cfd7-ff82-b23b-6334910518e4
ms.date: 06/08/2017
---


# Required client settings

**Applies to:** Access 2013 | Access 2016

Specify the following settings to use a custom **DataFactory** handler.

- Specify "Provider=MS Remote" in the **Connection** object **Provider** property or the **Connection** object connection string "**Provider=**" keyword.
    
- Set the **CursorLocation** property to **adUseClient**.
    
- Specify the name of the handler to use in the RDS.DataControl object's **Handler** property, or the **Recordset** object's connection string "**Handler=**" keyword. (You cannot set the handler in the **Connection** object connect string.)
    
RDS provides a default handler on the server named **MSDFMAP.Handler**. (The default customization file is named **MSDFMAP.INI**.)

## Example

Assume that the following sections in **MSDFMAP.INI** and the data source name, AdvWorks, have been previously defined:

```sql
[connect CustomerDataBase] 
Access=ReadWrite 
Connect="DSN=AdvWorks" 
 
[sql CustomerById] 
SQL="SELECT * FROM Customers WHERE CustomerID = ?" 

```

The following code snippets are written in Visual Basic.

### RDS.DataControl version

```vb
Dim dc as New RDS.DataControl 
Set dc.Handler = "MSDFMAP.Handler" 
Set dc.Server = "http://yourServer" 
Set dc.Connect = "Data Source=CustomerDatabase" 
Set dc.SQL = "CustomerById(4)" 
dc.Refresh
```


### Recordset version

```vb
Dim rs as New ADODB.Recordset 
rs.CursorLocation = adUseClient
```

Specify either the **Handler** property or keyword, the **Provider** property or keyword, and the _CustomerById_ and _CustomerDatabase_ identifiers. Then open the **Recordset** object.


```vb
rs.Open "CustomerById(4)", "Handler=MSDFMAP.Handler;" &; _ 
   "Provider=MS Remote;Data Source=CustomerDatabase;" &; _ 
   "Remote Server=http://yourServer" 

```

## See also

- [Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/en-us/msoffice/forum?page=1&;tab=question&;status=all&;auth=1)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)
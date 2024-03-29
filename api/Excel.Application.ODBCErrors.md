---
title: Application.ODBCErrors property (Excel)
keywords: vbaxl10.chm133174
f1_keywords:
- vbaxl10.chm133174
api_name:
- Excel.Application.ODBCErrors
ms.assetid: 47caef7a-fd3c-f67f-09c1-5ac21d65b67f
ms.date: 04/05/2019
ms.localizationpriority: medium
---


# Application.ODBCErrors property (Excel)

Returns an **[ODBCErrors](excel.odbcerrors.md)** collection that contains all the ODBC errors generated by the most recent query table or PivotTable report operation. Read-only.


## Syntax

_expression_.**ODBCErrors**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

If there's more than one query running at the same time, the **ODBCErrors** collection contains the ODBC errors from the query that finished last.


## Example

This example refreshes query table one and displays any ODBC errors that occur.

```vb
With Worksheets(1).QueryTables(1) 
 .Refresh 
 Set errs = Application.ODBCErrors 
 If errs.Count > 0 Then 
 Set r = .Destination.Cells(1) 
 r.Value = "The following errors occurred:" 
 c = 0 
 For Each er In errs 
 c = c + 1 
 r.offset(c, 0).value = er.ErrorString 
 r.offset(c, 1).value = er.SqlState 
 Next 
 Else 
 MsgBox "Query complete: all records returned." 
 End If 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
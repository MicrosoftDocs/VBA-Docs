---
title: WorkbookConnection.Refresh method (Excel)
keywords: vbaxl10.chm774081
f1_keywords:
- vbaxl10.chm774081
api_name:
- Excel.WorkbookConnection.Refresh
ms.assetid: 5e6f045f-6625-857c-eb55-ac52f70e8fb9
ms.date: 10/06/2026
ms.localizationpriority: medium
---


# WorkbookConnection.Refresh method (Excel)

Refreshes a workbook connection.


## Syntax

_expression_.**Refresh**

_expression_ A variable that represents a **[WorkbookConnection](Excel.WorkbookConnection.md)** object.


## Remarks

If the **[DisplayAlerts](Excel.Application.DisplayAlerts.md)** property is **False**, dialog boxes are not displayed, and the **Refresh** method fails with the Insufficient Connection Information exception.

A refresh failure for one connection will not have any impact on refresh operations for the other connections.

If the connection loads a query created with Power Query into a worksheet table, background refresh is turned off for the connection (the **[BackgroundQuery](Excel.OLEDBConnection.BackgroundQuery.md)** property of its **OLEDBConnection** object is **False**), and the query as a whole fails to evaluate (for example, because of an `error` expression in a step, or a function such as `Number.FromText` that fails while a step's value is computed rather than inside a single cell), the **Refresh** method raises run-time error 1004 with a generic description ("Application-defined or object-defined error"). The error message produced by the query isn't included. To get the query's error message, refresh the table by using the **[Refresh](Excel.QueryTable.Refresh.md)** method of its **[QueryTable](Excel.QueryTable.md)** object with background refresh turned off instead. That method raises the same error number, and its description includes the query's error, for example `[Expression.Error] ...`.

If background refresh is turned on, the **Refresh** method returns before the refresh completes, and errors that occur during the refresh aren't returned to the calling code.

Errors in individual cells of the query's result don't cause the refresh to fail. The refresh succeeds, and those cells are left empty.


## Example

This example refreshes every query table on the active sheet and writes the error message of each failed refresh to the Immediate window. Calling **Refresh** on the **WorkbookConnection** objects instead would report only the generic description.

```vb
Sub RefreshQueryTables()
    Dim lo As ListObject
    For Each lo In ActiveSheet.ListObjects
        If lo.SourceType = xlSrcQuery Then
            On Error Resume Next
            ' lo.QueryTable.WorkbookConnection.Refresh would report only a generic description.
            lo.QueryTable.Refresh BackgroundQuery:=False
            If Err.Number <> 0 Then Debug.Print lo.Name & ": " & Err.Description
            On Error GoTo 0
        End If
    Next lo
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
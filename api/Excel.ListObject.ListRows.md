---
title: ListObject.ListRows property (Excel)
keywords: vbaxl10.chm734087
f1_keywords:
- vbaxl10.chm734087
ms.prod: excel
api_name:
- Excel.ListObject.ListRows
ms.assetid: 7b584f41-ffc0-abe4-e755-ef163bcbb2ed
ms.date: 04/30/2019
localization_priority: Normal
---


# ListObject.ListRows property (Excel)

Returns a **[ListRows](Excel.ListRows.md)** object that represents all the rows of data in the **ListObject** object. Read-only.


## Syntax

_expression_.**ListRows**

_expression_ A variable that represents a **[ListObject](Excel.ListObject.md)** object.


## Remarks

The **ListRows** object returned does not include the header, total, or Insert rows.


## Example

The following example deletes a row specified by number in the **ListRows** collection that is created by a call to the **ListRows** property.

```vb
Sub DeleteListRow(iRowNumber As Integer) 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 Dim objListRows As ListRows 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 Set objListRows = objListObj.ListRows 
 
 If (iRowNumber <> 0) And (iRowNumber < objListRows.Count - 1) Then 
 objListRows(iRowNumber).Delete 
 End If 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

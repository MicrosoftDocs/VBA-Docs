---
title: ListObject.ListColumns property (Excel)
keywords: vbaxl10.chm734086
f1_keywords:
- vbaxl10.chm734086
ms.prod: excel
api_name:
- Excel.ListObject.ListColumns
ms.assetid: 64cefe01-b0e6-1cdd-3eec-7cb8389666dc
ms.date: 04/30/2019
localization_priority: Normal
---


# ListObject.ListColumns property (Excel)

Returns a **[ListColumns](Excel.ListColumns.md)** collection that represents all the columns in a **ListObject** object. Read-only.


## Syntax

_expression_.**ListColumns**

_expression_ A variable that represents a **[ListObject](Excel.ListObject.md)** object.


## Example

The following example displays the name of the second column in the **ListColumns** collection object as created by a call to the **ListColumns** property. For this code to run, the Sheet1 worksheet must contain a table.

```vb
Sub DisplayColumnName 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 Dim objListCols As ListColumns 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 Set objListCols = objListObj.ListColumns 
 
 Debug.Print objListCols(2).Name 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

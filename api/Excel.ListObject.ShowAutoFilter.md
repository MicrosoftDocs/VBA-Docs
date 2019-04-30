---
title: ListObject.ShowAutoFilter property (Excel)
keywords: vbaxl10.chm734091
f1_keywords:
- vbaxl10.chm734091
ms.prod: excel
api_name:
- Excel.ListObject.ShowAutoFilter
ms.assetid: ae9dfc8d-dd58-802d-2e96-461abdb9ee2b
ms.date: 04/30/2019
localization_priority: Normal
---


# ListObject.ShowAutoFilter property (Excel)

Returns **Boolean** to indicate whether the AutoFilter will be displayed. Read/write **Boolean**.


## Syntax

_expression_.**ShowAutoFilter**

_expression_ A variable that represents a **[ListObject](Excel.ListObject.md)** object.


## Remarks

The **ShowAutoFilter** property defaults to **True** for a new **ListObject** object.


## Example

The following example displays the setting of the **ShowAutoFilter** property of the default list on Sheet1 of the active workbook.

```vb
 
 Dim wrksht As Worksheet 
 Dim oListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set oListCol = wrksht.ListObjects(1) 
 
 Debug.Print oListCol.ShowAutoFilter
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
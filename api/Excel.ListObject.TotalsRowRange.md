---
title: ListObject.TotalsRowRange property (Excel)
keywords: vbaxl10.chm734094
f1_keywords:
- vbaxl10.chm734094
ms.prod: excel
api_name:
- Excel.ListObject.TotalsRowRange
ms.assetid: 80f22712-5113-30d9-a0ea-1158a563d17b
ms.date: 06/08/2017
localization_priority: Normal
---


# ListObject.TotalsRowRange property (Excel)

Returns a **[Range](Excel.Range(object).md)** object representing the Total row, if any, from a specified **ListObject** object. Read-only.


## Syntax

_expression_.**TotalsRowRange**

_expression_ A variable that represents a **[ListObject](Excel.ListObject.md)** object.


## Example

The following sample code returns the address of the Total row in the default list on Sheet1 of the active workbook. The code displays the Total row if it is not displayed already.

```vb
Sub DisplayTotalsRowAddress() 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet2") 
 Set objListObj = wrksht.ListObjects(1) 
 objListObj.ShowTotals = True 
 MsgBox objListObj.TotalsRowRange.Address 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: ListObject.ShowTotals property (Excel)
keywords: vbaxl10.chm734092
f1_keywords:
- vbaxl10.chm734092
ms.prod: excel
api_name:
- Excel.ListObject.ShowTotals
ms.assetid: 99a86f33-d718-98df-9869-76d52ddab0bb
ms.date: 04/30/2019
localization_priority: Normal
---


# ListObject.ShowTotals property (Excel)

Gets or sets a **Boolean** to indicate whether the Total row is visible. Read/write **Boolean**.


## Syntax

_expression_.**ShowTotals**

_expression_ A variable that represents a **[ListObject](Excel.ListObject.md)** object.


## Example

The following code sample displays the current setting of the **ShowTotals** property of the default list on Sheet1 of the active workbook.

```vb
 
Sub test() 
Dim oListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set oListObj = wrksht.ListObjects(1) 
 
 Debug.Print oListObj.ShowTotals 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
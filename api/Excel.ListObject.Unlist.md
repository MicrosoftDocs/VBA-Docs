---
title: ListObject.Unlist method (Excel)
keywords: vbaxl10.chm734077
f1_keywords:
- vbaxl10.chm734077
ms.prod: excel
api_name:
- Excel.ListObject.Unlist
ms.assetid: 030f8f78-08e1-8a49-ee06-a7b4254aa5fc
ms.date: 04/30/2019
localization_priority: Normal
---


# ListObject.Unlist method (Excel)

Removes the list functionality from a **ListObject** object. After you use this method, the range of cells that made up the list will be a regular range of data.


## Syntax

_expression_.**Unlist**

_expression_ A variable that represents a **[ListObject](Excel.ListObject.md)** object.


## Remarks

Running this method leaves the cell data, formatting, and formulas on the worksheet. The Total row is also left intact. This method removes any link to a Microsoft SharePoint Foundation site. **AutoFilter** is also removed from the list.


## Example

The following example removes the list features from a list on a worksheet.

```vb
Sub DeList() 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 objListObj.Unlist 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
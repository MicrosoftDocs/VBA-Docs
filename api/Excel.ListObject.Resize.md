---
title: ListObject.Resize method (Excel)
keywords: vbaxl10.chm734079
f1_keywords:
- vbaxl10.chm734079
ms.prod: excel
api_name:
- Excel.ListObject.Resize
ms.assetid: b9a0ae05-d1cd-3ce6-f4ae-6a539850a1b5
ms.date: 06/08/2019
localization_priority: Normal
---


# ListObject.Resize method (Excel)

The **Resize** method allows a **ListObject** object to be resized over a new range. No cells are inserted or moved.


## Syntax

_expression_.**Resize** (_Range_)

_expression_ An expression that returns a **[ListObject](Excel.ListObject.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **[Range](Excel.Range(object).md)**|The new range.|

## Remarks

For tables that are linked to a server that is running Microsoft SharePoint Foundation, you can resize the list using this method by providing a _Range_ argument that differs from the current range of the **ListObject** only in the number of rows it contains. Attempting to resize lists linked to SharePoint Foundation by adding or deleting columns (in the _Range_ argument) results in a run-time error.


## Example

The following example uses the **Resize** method to resize the default **ListObject** object on Sheet1 of the active workbook.

```vb
Sub ResizeList() 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 
 objListObj.Resize Range("A1:B10") 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

---
title: Chart.OLEObjects method (Excel)
keywords: vbaxl10.chm149126
f1_keywords:
- vbaxl10.chm149126
ms.prod: excel
api_name:
- Excel.Chart.OLEObjects
ms.assetid: e42150c1-8661-75b4-f1e8-fec8cc82f59b
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.OLEObjects method (Excel)

Returns an object that represents either a single OLE object (an **[OLEObject](Excel.OLEObject.md)**) or a collection of all OLE objects (an **[OLEObjects](Excel.OLEObjects.md)** collection) on the chart or sheet. Read-only.


## Syntax

_expression_.**OLEObjects** (_Index_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the OLE object.|

## Return value

Object


## Example

This example creates a list of link types for OLE objects on Sheet1. The list appears on a new worksheet created by the example.

```vb
Set newSheet = Worksheets.Add 
i = 2 
newSheet.Range("A1").Value = "Name" 
newSheet.Range("B1").Value = "Link Type" 
For Each obj In Worksheets("Sheet1").OLEObjects 
 newSheet.Cells(i, 1).Value = obj.Name 
 If obj.OLEType = xlOLELink Then 
 newSheet.Cells(i, 2) = "Linked" 
 Else 
 newSheet.Cells(i, 2) = "Embedded" 
 End If 
 i = i + 1 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
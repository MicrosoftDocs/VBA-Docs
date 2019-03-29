---
title: PivotCaches object (Excel)
keywords: vbaxl10.chm228072
f1_keywords:
- vbaxl10.chm228072
ms.prod: excel
api_name:
- Excel.PivotCaches
ms.assetid: cfd979b9-d52f-f34b-4b66-4fb17efcdc92
ms.date: 03/30/2019
localization_priority: Normal
---


# PivotCaches object (Excel)

Represents the collection of memory caches from the PivotTable reports in a workbook.


## Remarks

Each memory cache is represented by a **[PivotCache](Excel.PivotCache.md)** object.


## Example

Use the **[PivotCaches](Excel.Workbook.PivotCaches.md)** method of the **Workbook** object to return the **PivotCaches** collection. 

The following example sets the **[RefreshOnFileOpen](Excel.PivotCache.RefreshOnFileOpen.md)** property for all memory caches in the active workbook.


```vb
For Each pc In ActiveWorkbook.PivotCaches 
 pc.RefreshOnFileOpen = True 
Next
```


## Methods

- [Create](Excel.PivotCaches.Create.md)
- [Item](Excel.PivotCaches.Item.md)

## Properties

- [Application](Excel.PivotCaches.Application.md)
- [Count](Excel.PivotCaches.Count.md)
- [Creator](Excel.PivotCaches.Creator.md)
- [Parent](Excel.PivotCaches.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
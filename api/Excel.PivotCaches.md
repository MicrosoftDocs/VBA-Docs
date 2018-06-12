---
title: PivotCaches Object (Excel)
keywords: vbaxl10.chm228072
f1_keywords:
- vbaxl10.chm228072
ms.prod: excel
api_name:
- Excel.PivotCaches
ms.assetid: cfd979b9-d52f-f34b-4b66-4fb17efcdc92
ms.date: 06/08/2017
---


# PivotCaches Object (Excel)

Represents the collection of memory caches from the PivotTable reports in a workbook.


## Remarks

 Each memory cache is represented by a **[PivotCache](Excel.PivotCache.md)** object.


## Example

Use the  **[PivotCaches](Excel.Workbook.PivotCaches.md)** method to return the **[PivotCaches](Excel.PivotCaches.md)** collection. The following example sets the **[RefreshOnFileOpen](Excel.PivotCache.RefreshOnFileOpen.md)** property for all memory caches in the active workbook.


```
For Each pc In ActiveWorkbook.PivotCaches 
 pc.RefreshOnFileOpen = True 
Next
```


## Methods



|**Name**|
|:-----|
|[Create](Excel.PivotCaches.Create.md)|
|[Item](Excel.PivotCaches.Item.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.PivotCaches.Application.md)|
|[Count](Excel.PivotCaches.Count.md)|
|[Creator](Excel.PivotCaches.Creator.md)|
|[Parent](Excel.PivotCaches.Parent.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

---
title: SlicerItem.HasData property (Excel)
keywords: vbaxl10.chm907080
f1_keywords:
- vbaxl10.chm907080
ms.prod: excel
api_name:
- Excel.SlicerItem.HasData
ms.assetid: 17ce0cdc-ec30-638a-e869-4640ee0ef5a3
ms.date: 05/16/2019
localization_priority: Normal
---


# SlicerItem.HasData property (Excel)

Returns whether the slicer item contains data that matches the current manual filter state. Read-only.


## Syntax

_expression_.**HasData**

_expression_ A variable that represents a **[SlicerItem](Excel.SlicerItem.md)** object.


## Return value

Boolean


## Remarks

**True** if the slicer item contains data that matches the current manual filter state; otherwise, **False**. This property is only supported when cross filtering is turned on, and will generate a run-time error otherwise. 

To determine if cross filtering is turned on for a slicer associated with a PivotTable report, check the value of the **[CrossFilterType](Excel.SlicerCache.CrossFilterType.md)** property of the specified **SlicerItem** object's parent **SlicerCache** object. 

To determine if cross filtering is turned on for a slicer associated with an OLAP data source, check the value of the **[CrossFilterType](Excel.SlicerCacheLevel.CrossFilterType.md)** property of the **SlicerCacheLevel** object that corresponds to the OLAP hierarchy being filtered.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
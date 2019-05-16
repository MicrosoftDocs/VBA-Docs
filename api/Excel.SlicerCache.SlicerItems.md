---
title: SlicerCache.SlicerItems property (Excel)
keywords: vbaxl10.chm897083
f1_keywords:
- vbaxl10.chm897083
ms.prod: excel
api_name:
- Excel.SlicerCache.SlicerItems
ms.assetid: d552a519-3d9f-74b8-4cbe-3b5c935a14d9
ms.date: 05/16/2019
localization_priority: Normal
---


# SlicerCache.SlicerItems property (Excel)

Returns a **[SlicerItems](Excel.SlicerItems.md)** collection that contains the collection of all items in the slicer cache. Read-only.


## Syntax

_expression_.**SlicerItems**

_expression_ A variable that represents a **[SlicerCache](Excel.SlicerCache.md)** object.


## Return value

**SlicerItems**


## Remarks

The **SlicerItems** property is only applicable for slicers that are based on PivotTables based on workbook ranges or lists (**SlicerCache**.**SourceType** = **xlDatabase**), or for slicers that are based on PivotTables based on relational data sources (**SlicerCache**.**SourceType** = **xlExternal** and **SlicerCache**.**[OLAP](Excel.SlicerCache.OLAP.md)** = **False**). 

Attempting to access the **SlicerItems** property for slicers that are connected to an external OLAP data source (**SlicerCache**.**OLAP** = **True**) generates a run-time error. For OLAP data sources, use the **[SlicerItems](Excel.SlicerCacheLevel.SlicerItems.md)** property of the **SlicerCacheLevel** object instead.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: SlicerCache.VisibleSlicerItems property (Excel)
keywords: vbaxl10.chm897081
f1_keywords:
- vbaxl10.chm897081
ms.prod: excel
api_name:
- Excel.SlicerCache.VisibleSlicerItems
ms.assetid: ea9d1b43-1280-5423-515f-8d00e0624901
ms.date: 05/16/2019
localization_priority: Normal
---


# SlicerCache.VisibleSlicerItems property (Excel)

Returns a **[SlicerItems](Excel.SlicerItems.md)** collection that contains the collection of all the visible items in the specified slicer cache. Read-only.


## Syntax

_expression_.**VisibleSlicerItems**

_expression_ A variable that represents a **[SlicerCache](Excel.SlicerCache.md)** object.


## Return value

**SlicerItems**


## Remarks

The **VisibleSlicerItems** property is only applicable for slicers that are based on PivotTables based on workbook ranges or lists (**SlicerCache**.**SourceType** = **xlDatabase**). 

Attempting to access the **VisibleSlicerItems** property for slicers that are connected to an OLAP data source (**SlicerCache**.**OLAP** = **True**) generates a run-time error.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
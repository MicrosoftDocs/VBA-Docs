---
title: SlicerCacheLevel.VisibleSlicerItemsList property (Excel)
keywords: vbaxl10.chm901079
f1_keywords:
- vbaxl10.chm901079
ms.prod: excel
api_name:
- Excel.SlicerCacheLevel.VisibleSlicerItemsList
ms.assetid: 68c0800b-4130-59f2-d0c0-7cad49b98f0d
ms.date: 05/16/2019
localization_priority: Normal
---


# SlicerCacheLevel.VisibleSlicerItemsList property (Excel)

Returns the list of slicer items that are currently included in the slicer filter. Read-only.


## Syntax

_expression_.**VisibleSlicerItemsList**

_expression_ A variable that represents a **[SlicerCacheLevel](Excel.SlicerCacheLevel.md)** object.


## Return value

**Variant**


## Remarks

The list of slicer items are returned as MDX unique name strings. If this list is empty, the slicer is not filtering the data source and all slicer tiles are displayed as selected.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
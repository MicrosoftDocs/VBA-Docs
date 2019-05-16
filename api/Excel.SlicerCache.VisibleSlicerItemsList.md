---
title: SlicerCache.VisibleSlicerItemsList property (Excel)
keywords: vbaxl10.chm897082
f1_keywords:
- vbaxl10.chm897082
ms.prod: excel
api_name:
- Excel.SlicerCache.VisibleSlicerItemsList
ms.assetid: 1002d654-8207-fe88-567e-8bd4e36fbeb4
ms.date: 05/16/2019
localization_priority: Normal
---


# SlicerCache.VisibleSlicerItemsList property (Excel)

Returns or sets the list of MDX unique names for members at all levels of the hierarchy where manual filtering is applied. Read/write.


## Syntax

_expression_.**VisibleSlicerItemsList**

_expression_ A variable that represents a **[SlicerCache](Excel.SlicerCache.md)** object.


## Return value

**Variant**


## Remarks

The **VisibleSlicerItemsList** property is only applicable for slicers that are based on OLAP data sources (**SlicerCache**.**[OLAP](Excel.SlicerCache.OLAP.md)** = **True**).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
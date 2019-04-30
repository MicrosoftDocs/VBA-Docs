---
title: SlicerCache.SourceName property (Excel)
keywords: vbaxl10.chm897086
f1_keywords:
- vbaxl10.chm897086
ms.prod: excel
api_name:
- Excel.SlicerCache.SourceName
ms.assetid: 659a7670-024e-3763-7d94-e2e4b86cfc9e
ms.date: 06/08/2017
localization_priority: Normal
---


# SlicerCache.SourceName property (Excel)

Returns the name of the data source the slicer is connected to. Read-only


## Syntax

_expression_. `SourceName`

_expression_ A variable that represents a **[SlicerCache](Excel.SlicerCache.md)** object.


## Return value

 **String**


## Remarks

For slicers based on data in the workbook (**SlicerCache**. **[SourceType](Excel.SlicerCache.SourceType.md)** = **xlDatabase**), or slicers based on non-OLAP external data (**SlicerCache**. **SourceType** = **xlExternal** and **SlicerCache**. **[OLAP](Excel.SlicerCache.OLAP.md)** = **False**), returns the name of the corresponding column in the source data.

For OLAP slicers (**SlicerCache**. **OLAP** = **True**), returns the MDX unique name of the hierarchy the slicer is based on.


## See also


[SlicerCache Object](Excel.SlicerCache.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
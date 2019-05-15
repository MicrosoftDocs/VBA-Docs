---
title: SlicerCache.Name property (Excel)
keywords: vbaxl10.chm897080
f1_keywords:
- vbaxl10.chm897080
ms.prod: excel
api_name:
- Excel.SlicerCache.Name
ms.assetid: 3b4a00c0-c6c9-6eee-043c-8102642354df
ms.date: 05/16/2019
localization_priority: Normal
---


# SlicerCache.Name property (Excel)

Returns or sets the name of the slicer cache.


## Syntax

_expression_.**Name**

_expression_ A variable that returns a **[SlicerCache](Excel.SlicerCache.md)** object.


## Return value

**String**


## Remarks

The name of the slicer cache must be unique within the workbook namespace. By default, the name assigned to a slicer cache is  `Slicer_` followed by the name of the PivotTable field that the slicer cache is associated with. 

For example, if slicer is associated with the Product Category field in the PivotTable, the default name will be `Slicer_Product_Category` (any spaces in the field name are replaced with underscore characters). If there is more than one Product Category field in the same workbook with a slicer associated with it, or some other named entity in the workbook with the name `Slicer_Product_Category`, Excel appends a number after the name to produce a unique name, such as  `Slicer_Product_Category1`.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: SlicerCache.SortUsingCustomLists property (Excel)
keywords: vbaxl10.chm897087
f1_keywords:
- vbaxl10.chm897087
ms.prod: excel
api_name:
- Excel.SlicerCache.SortUsingCustomLists
ms.assetid: 61c156fe-67cf-f6e8-4fce-bc617c9a1e03
ms.date: 05/16/2019
localization_priority: Normal
---


# SlicerCache.SortUsingCustomLists property (Excel)

Returns or sets whether items in the specified slicer cache will be sorted by the custom lists. Read/write.


## Syntax

_expression_.**SortUsingCustomLists**

_expression_ A variable that represents a **[SlicerCache](Excel.SlicerCache.md)** object.


## Remarks

The **SortUsingCustomLists** property corresponds to the setting of the **Use Custom Lists when sorting** check box of the **Slicer Settings** dialog box. To access the custom lists associated with the current installation of Excel, choose the **File** tab > **Options** > **Advanced**, and then choose **Edit Custom Lists** under the **General** category.

The **SortUsingCustomLists** property only applies to slicers that are filtering non-OLAP data sources. Attempting to access this property from a slicer cache that is filtering an OLAP data source (**SlicerCache**.**[OLAP](Excel.SlicerCache.OLAP.md)** = **True**) generates a run-time error.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
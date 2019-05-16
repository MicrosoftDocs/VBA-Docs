---
title: SlicerCache.CrossFilterType property (Excel)
keywords: vbaxl10.chm897084
f1_keywords:
- vbaxl10.chm897084
ms.prod: excel
api_name:
- Excel.SlicerCache.CrossFilterType
ms.assetid: 8a29b376-c999-472d-0853-2e2f4a0949a0
ms.date: 05/16/2019
localization_priority: Normal
---


# SlicerCache.CrossFilterType property (Excel)

Returns or sets whether a slicer is participating in cross filtering with other slicers that share the same slicer cache, and how cross filtering is displayed. Read/write.


## Syntax

_expression_.**CrossFilterType**

_expression_ A variable that represents a **[SlicerCache](Excel.SlicerCache.md)** object.


## Return value

**[XlSlicerCrossFilterType](Excel.XlSlicerCrossFilterType.md)**


## Remarks

If more than one slicer is associated with the same PivotTable, by default, if the item or items that you filter by in one slicer have no corresponding data in another slicer, those items will be grayed out. For example, if you have a Country slicer and a State slicer, and you choose a country in the Country slicer, all states that are not in that country will be grayed out. This feature is referred to as *cross filtering*. 

The user interface settings that correspond to the setting of the **CrossFilterType** property are the **Visually indicate items with no data** and **Show items with no data last** check boxes in the **Slicer Settings** dialog box. 

Setting the **CrossFilterType** property to **xlSlicerCrossFilterShowItemsWithDataAtTop** corresponds to selecting both the **Visually indicate items with no data** and **Show items with no data last** check boxes. 

Setting the **CrossFilterType** property to **xlSlicerCrossFilterShowItemsWithNoData** corresponds to selecting only the **Visually indicate items with no data** check box. 

Clearing both check boxes corresponds to setting the **CrossFilterType** property to **xlSlicerNoCrossFilter**.

OLAP data sources (**SlicerCache**.**OLAP** = **True**) are not supported by the **CrossFilterType** property. For OLAP data sources, use the **[CrossFilterType](Excel.SlicerCacheLevel.CrossFilterType.md)** property of the **SlicerCacheLevel** object instead.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
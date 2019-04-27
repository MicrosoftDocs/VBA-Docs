---
title: IconSets.Item property (Excel)
keywords: vbaxl10.chm820076
f1_keywords:
- vbaxl10.chm820076
ms.prod: excel
api_name:
- Excel.IconSets.Item
ms.assetid: 79c0d577-f988-31c1-7a29-95f5d924cbc4
ms.date: 04/27/2019
localization_priority: Normal
---


# IconSets.Item property (Excel)

Returns a single **[IconSet](Excel.IconSet.md)** object from the **IconSets** collection. Read-only.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents an **[IconSets](Excel.IconSets.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number of the **IconSet** object.|


## Remarks

The value of the _Index_ parameter cannot be greater than the number of icon sets available. To find the number of icon sets available to the workbook, use the **[Count](excel.iconsets.count.md)** property.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
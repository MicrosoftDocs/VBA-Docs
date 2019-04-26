---
title: IconSetCondition.ModifyAppliesToRange method (Excel)
keywords: vbaxl10.chm812077
f1_keywords:
- vbaxl10.chm812077
ms.prod: excel
api_name:
- Excel.IconSetCondition.ModifyAppliesToRange
ms.assetid: 89f13781-2641-4cba-061b-cc82b479b6df
ms.date: 04/27/2019
localization_priority: Normal
---


# IconSetCondition.ModifyAppliesToRange method (Excel)

Sets the cell range to which this formatting rule applies.


## Syntax

_expression_.**ModifyAppliesToRange** (_Range_)

_expression_ A variable that represents an **[IconSetCondition](Excel.IconSetCondition.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|The range to which this formatting rule will be applied.|

## Remarks

The range must be in the A1 reference style and be entirely contained within the sheet that is the parent of the **[FormatConditions](Excel.FormatConditions.md)** collection. It can include the range operator (a colon), the intersection operator (a space), or the union operator (a comma). Dollar signs can also be used, but they are ignored.

You can also use a local defined name in any part of the range, but the name must be in the language of the macro.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
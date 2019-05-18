---
title: Top10.ModifyAppliesToRange method (Excel)
keywords: vbaxl10.chm822087
f1_keywords:
- vbaxl10.chm822087
ms.prod: excel
api_name:
- Excel.Top10.ModifyAppliesToRange
ms.assetid: 3baf8e16-4bb7-ec97-da0a-17187500f1f1
ms.date: 05/18/2019
localization_priority: Normal
---


# Top10.ModifyAppliesToRange method (Excel)

Sets the cell range to which this formatting rule applies.


## Syntax

_expression_.**ModifyAppliesToRange** (_Range_)

_expression_ A variable that represents a **[Top10](Excel.Top10.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|The range to which this formatting rule will be applied.|

## Remarks

The range must be in the A1 reference style and be entirely contained within the sheet that is the parent of the **[FormatConditions](Excel.FormatConditions.md)** collection. It can include the range operator (a colon), the intersection operator (a space), or the union operator (a comma). Dollar signs can also be used, but they are ignored.

You can also use a local defined name in any part of the range, but the name must be in the language of the macro.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
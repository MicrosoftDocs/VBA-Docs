---
title: Application.SheetBeforeRightClick event (Excel)
keywords: vbaxl10.chm504076
f1_keywords:
- vbaxl10.chm504076
ms.prod: excel
api_name:
- Excel.Application.SheetBeforeRightClick
ms.assetid: eb91ede3-3f17-7cf8-2b6f-b519acd11ce3
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.SheetBeforeRightClick event (Excel)

Occurs when any worksheet is right-clicked, before the default right-click action.


## Syntax

_expression_.**SheetBeforeRightClick** (_Sh_, _Target_, _Cancel_)

_expression_ An expression that returns an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|A **[Worksheet](Excel.Worksheet.md)** object that represents the sheet.|
| _Target_|Required| **Range**|The cell nearest to the mouse pointer when the right-click occurred.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the default right-click action isn't performed when the procedure is finished.|

## Remarks

This event doesn't occur on chart sheets.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
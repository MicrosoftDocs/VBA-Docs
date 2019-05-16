---
title: Sheets.Add2 method (Excel)
keywords: vbaxl10.chm152090
f1_keywords:
- vbaxl10.chm152090
ms.prod: excel
ms.assetid: f44b3ef1-8452-4e26-b91c-d24124fa5bc6
ms.date: 05/15/2019
localization_priority: Normal
---


# Sheets.Add2 method (Excel)

This method is only implemented for the **[Charts](excel.charts.md)** collection object and will produce a run-time error if used on the **Sheets** and **[Worksheets](Excel.Worksheets.md)** objects.


## Syntax

_expression_.**Add2** (_Before_, _After_, _Count_, _NewLayout_)

_expression_ A variable that represents a **[Sheets](Excel.Sheets.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Before_|Optional|**Variant**|An object that specifies the sheet before which the new sheet is added.|
| _After_|Optional|**Variant**|An object that specifies the sheet after which the new sheet is added.|
| _Count_|Optional|**Variant**|The number of sheets to be added. The default value is one.|
| _NewLayout_|Optional|**Variant**|The layout of the new worksheet.|

## Return value

**OBJECT**


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

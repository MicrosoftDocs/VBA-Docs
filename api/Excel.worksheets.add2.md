---
title: Worksheets.Add2 method (Excel)
keywords: vbaxl10.chm470090
f1_keywords:
- vbaxl10.chm470090
ms.prod: excel
ms.assetid: 4ae91335-f714-45e4-9677-6dfece31342e
ms.date: 05/18/2019
localization_priority: Normal
---


# Worksheets.Add2 method (Excel)

This method is only implemented for the **[Charts](excel.charts.md)** collection object and will produce a run-time error if used on the **[Sheets](Excel.Sheets.md)** and **Worksheets** objects.


## Syntax

_expression_.**Add2** (_Before_, _After_, _Count_, _NewLayout_)

_expression_ A variable that represents a **[Worksheets](Excel.Worksheets.md)** object.


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
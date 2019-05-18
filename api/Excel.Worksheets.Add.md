---
title: Worksheets.Add method (Excel)
keywords: vbaxl10.chm470073
f1_keywords:
- vbaxl10.chm470073
ms.prod: excel
api_name:
- Excel.Worksheets.Add
ms.assetid: c771d87a-64e1-e292-9db4-54386a69301e
ms.date: 05/18/2019
localization_priority: Normal
---


# Worksheets.Add method (Excel)

Creates a new worksheet, chart, or macro sheet. The new worksheet becomes the active sheet.


## Syntax

_expression_.**Add** (_Before_, _After_, _Count_, _Type_)

_expression_ A variable that represents a **[Worksheets](Excel.Worksheets.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Before_|Optional| **Variant**|An object that specifies the sheet before which the new sheet is added.|
| _After_|Optional| **Variant**|An object that specifies the sheet after which the new sheet is added.|
| _Count_|Optional| **Variant**|The number of sheets to be added. The default value is one.|
| _Type_|Optional| **Variant**|Specifies the sheet type. Can be one of the following **[XlSheetType](Excel.XlSheetType.md)** constants: **xlWorksheet**, **xlChart**, **xlExcel4MacroSheet**, or **xlExcel4IntlMacroSheet**. If you are inserting a sheet based on an existing template, specify the path to the template. The default value is **xlWorksheet**.|

## Return value

An Object value that represents the new worksheet, chart, or macro sheet.


## Remarks

If _Before_ and _After_ are both omitted, the new sheet is inserted before the active sheet.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

---
title: Sheets.Add method (Excel)
keywords: vbaxl10.chm152073
f1_keywords:
- vbaxl10.chm152073
ms.prod: excel
api_name:
- Excel.Sheets.Add
ms.assetid: db5de750-fd09-2b18-c52b-98d88eeb0ffc
ms.date: 09/03/2019
localization_priority: Normal
---


# Sheets.Add method (Excel)

Creates a new worksheet, chart, or macro sheet. The new worksheet becomes the active sheet.


## Syntax

_expression_.**Add** (_Before_, _After_, _Count_, _Type_)

_expression_ A variable that represents a **[Sheets](Excel.Sheets.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Before_|Optional| **Variant**|An object that specifies the sheet before which the new sheet is added.|
| _After_|Optional| **Variant**|An object that specifies the sheet after which the new sheet is added.|
| _Count_|Optional| **Variant**|The number of sheets to be added. The default value is the number of selected sheets.|
| _Type_|Optional| **Variant**|Specifies the sheet type. Can be one of the following **[XlSheetType](Excel.XlSheetType.md)** constants: **xlWorksheet**, **xlChart**, **xlExcel4MacroSheet**, or **xlExcel4IntlMacroSheet**. If you are inserting a sheet based on an existing template, specify the path to the template. The default value is **xlWorksheet**.|

## Return value

An Object value that represents the new worksheet, chart, or macro sheet.

## Remarks

If _Before_ and _After_ are both omitted, the new sheet is inserted before the active sheet.

## Example

This example inserts a new worksheet before the last worksheet in the active workbook.

```vb
ActiveWorkbook.Sheets.Add Before:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
```

<br/>

This example inserts a new worksheet after the last worksheet in the active workbook, and captures the returned object reference in a local variable.

```vb
Dim sheet As Worksheet
Set sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
```

> [!NOTE] 
> In 32-bit Excel 2010, this method cannot create more than 255 sheets at one time.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

---
title: Sheets.Move method (Excel)
keywords: vbaxl10.chm152079
f1_keywords:
- vbaxl10.chm152079
api_name:
- Excel.Sheets.Move
ms.assetid: 8cfb8888-b676-15ba-47eb-9d3d4dae5416
ms.date: 05/15/2019
ms.localizationpriority: medium
---


# Sheets.Move method (Excel)

Moves the sheet to another location in the workbook.


## Syntax

_expression_.**Move** (_Before_, _After_)

_expression_ A variable that represents a **[Sheets](Excel.Sheets.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Before_|Optional| **Variant**|The sheet before which the moved sheet will be placed. You cannot specify _Before_ if you specify _After_.|
| _After_|Optional| **Variant**| The sheet after which the moved sheet will be placed. You cannot specify _After_ if you specify _Before_.|

## Remarks

If you don't specify either _Before_ or _After_, Microsoft Excel creates a new workbook that contains the moved sheet.


## Example

This example moves Sheet1 after Sheet3 in the active workbook.

```vb
Worksheets("Sheet1").Move _ 
 after:=Worksheets("Sheet3")
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
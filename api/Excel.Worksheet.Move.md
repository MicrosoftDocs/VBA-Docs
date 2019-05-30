---
title: Worksheet.Move method (Excel)
keywords: vbaxl10.chm174079
f1_keywords:
- vbaxl10.chm174079
ms.prod: excel
api_name:
- Excel.Worksheet.Move
ms.assetid: 808e6eb8-7811-6f72-5acc-b3779587aa52
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.Move method (Excel)

Moves the sheet to another location in the workbook.


## Syntax

_expression_.**Move** (_Before_, _After_)

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


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

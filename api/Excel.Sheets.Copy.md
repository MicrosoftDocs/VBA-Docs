---
title: Sheets.Copy method (Excel)
keywords: vbaxl10.chm152074
f1_keywords:
- vbaxl10.chm152074
ms.prod: excel
api_name:
- Excel.Sheets.Copy
ms.assetid: 8cfee52e-dc0f-a54f-21ba-00a65ba2029c
ms.date: 06/08/2017
localization_priority: Priority
---


# Sheets.Copy method (Excel)

Copies the sheet to another location in the workbook.


## Syntax

_expression_. `Copy`( `_Before_` , `_After_` )

_expression_ A variable that represents a [Sheets](./Excel.Sheets.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Before_|Optional| **Variant**|The sheet before which the copied sheet will be placed. You cannot specify  _Before_ if you specify _After_.|
| _After_|Optional| **Variant**|The sheet after which the copied sheet will be placed. You cannot specify  _After_ if you specify _Before_.|

## Remarks

If you don't specify either  _Before_ or _After_, Microsoft Excel creates a new workbook that contains the copied sheet.


## Example

This example copies Sheet1, placing the copy after Sheet3.


```vb
Worksheets("Sheet1").Copy After:=Worksheets("Sheet3")
```


## See also


[Sheets Object](Excel.Sheets.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
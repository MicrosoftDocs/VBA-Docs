---
title: WorksheetFunction.SumIf method (Excel)
keywords: vbaxl10.chm137241
f1_keywords:
- vbaxl10.chm137241
ms.prod: excel
api_name:
- Excel.WorksheetFunction.SumIf
ms.assetid: 2df06641-0307-339f-236e-674d0bf58a78
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.SumIf method (Excel)

Adds the cells specified by a given criteria.


## Syntax

_expression_.**SumIf** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Range**|Range - the range of cells that you want evaluated by criteria.|
| _Arg2_|Required| **Variant**|Criteria - the criteria in the form of a number, expression, or text that defines which cells will be added. For example, criteria can be expressed as 32, "32", ">32", or "apples".|
| _Arg3_|Optional| **Variant**|Sum_range - the actual cells to add if their corresponding cells in range match criteria. If sum_range is omitted, the cells in range are both evaluated by criteria and added if they match criteria.|

## Return value

**Double**


## Remarks

Sum_range does not have to be the same size and shape as range. The actual cells that are added are determined by using the top, left cell in sum_range as the beginning cell, and then including cells that correspond in size and shape to range. For example:
    
|If range is|And sum_range is|The actual cells are|
|:-----|:-----|:-----|
|A1:A5|B1:B5|B1:B5|
|A1:A5|B1:B3|B1:B5|
|A1:B4|C1:D4|C1:D4|
|A1:B4|C1:C2|C1:D4|

You can use the wildcard characters, question mark (?) and asterisk (*), in criteria. A question mark matches any single character; an asterisk matches any sequence of characters. If you want to find an actual question mark or asterisk, type a tilde (~) preceding the character.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

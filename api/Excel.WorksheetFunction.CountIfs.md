---
title: WorksheetFunction.CountIfs method (Excel)
keywords: vbaxl10.chm137354
f1_keywords:
- vbaxl10.chm137354
ms.prod: excel
api_name:
- Excel.WorksheetFunction.CountIfs
ms.assetid: 399dcc8e-2523-8aa5-8112-b4cbc572d34e
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.CountIfs method (Excel)

Counts the number of cells within a range that meet multiple criteria.


## Syntax

_expression_.**CountIfs** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_, _Arg7_, _Arg8_, _Arg9_, _Arg10_, _Arg11_, _Arg12_, _Arg13_, _Arg14_, _Arg15_, _Arg16_, _Arg17_, _Arg18_, _Arg19_, _Arg20_, _Arg21_, _Arg22_, _Arg23_, _Arg24_, _Arg25_, _Arg26_, _Arg27_, _Arg28_, _Arg29_, _Arg30_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Range**|One or more ranges in which to evaluate the associated criteria.|
| _Arg2_ &ndash; _Arg30_|Required| **Variant**|One or more criteria in the form of a number, expression, cell reference, or text that define which cells will be counted. For example, criteria can be expressed as 32, "32", ">32", "apples", or B4.|

## Return value

**Double**


## Remarks

Each cell in a range is counted only if all of the corresponding criteria specified are **True** for that cell.
    
If a cell in any argument is an empty cell, **CountIfs** treats it as a 0 value.
    
You can use the wildcard characters, question mark (?) and asterisk (*), in criteria. A question mark matches any single character; an asterisk matches any sequence of characters. If you want to find an actual question mark or asterisk, type a tilde (~) before the character.
    
## See also

- [COUNTIFS function, Office Support](https://support.office.com/article/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

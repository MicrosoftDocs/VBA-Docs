---
title: WorksheetFunction.Fixed method (Excel)
keywords: vbaxl10.chm137084
f1_keywords:
- vbaxl10.chm137084
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Fixed
ms.assetid: befc65b2-0216-dbd7-e376-edbcbfe532c5
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Fixed method (Excel)

Rounds a number to the specified number of decimals, formats the number in decimal format using a period and commas, and returns the result as text.


## Syntax

_expression_.**Fixed** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - the number you want to round and convert to text.|
| _Arg2_|Optional| **Variant**|Decimals - the number of digits to the right of the decimal point.|
| _Arg3_|Optional| **Variant**|No_commas - a logical value that, if **True**, prevents **Fixed** from including commas in the returned text.|

## Return value

**String**


## Remarks

Numbers in Microsoft Excel can never have more than 15 significant digits, but decimals can be as large as 127.
    
If decimals is negative, number is rounded to the left of the decimal point.
    
If you omit decimals, it is assumed to be 2.
    
If no_commas is **False** or omitted, the returned text includes commas as usual.
    
The major difference between formatting a cell containing a number with the  **Cells** command (**Format** menu) and formatting a number directly with the **Fixed** function is that **Fixed** converts its result to text. A number formatted with the **Cells** command is still a number.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
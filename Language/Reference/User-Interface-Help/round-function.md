---
title: Round function (Visual Basic for Applications)
keywords: vblr6.chm1009020
f1_keywords:
- vblr6.chm1009020
ms.prod: office
ms.assetid: 897563a8-e66a-1ff1-36b2-da44ae56f48c
ms.date: 12/13/2018
localization_priority: Normal
---


# Round function

Returns a number rounded to a specified number of decimal places.

## Syntax

**Round**(_expression_, [ _numdecimalplaces_ ])

<br/>

The **Round** function syntax has these parts:

|Part|Description|
|:-----|:-----|
| _expression_|Required. [Numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) being rounded.|
| _numdecimalplaces_|Optional. Number indicating how many places to the right of the decimal are included in the rounding. If omitted, integers are returned by the **Round** function.|

> [!NOTE] 
> This VBA function returns something commonly referred to as bankers rounding. So be careful before using this function. For more predictable results, use [Worksheet Round](../../../api/excel.worksheetfunction.round.md) functions in Excel VBA.

## Example

```vb
?Round(0.12335,4)
 0,1234
?Round(0.12345,4)
 0,1234
?Round(0.12355,4)
 0,1236
?Round(0.12365,4)
 0,1236

?WorksheetFunction.Round(0.12345,4)
 0,1235
?WorksheetFunction.RoundUp(0.12345,4)
 0,1235
?WorksheetFunction.RoundDown(0.12345,4)
 0,1234

?Round(0.00005,4)
 0
?WorksheetFunction.Round(0.00005,4)
 0,0001
?WorksheetFunction.RoundUp(0.00005,4)
 0,0001
?WorksheetFunction.RoundDown(0.00005,4)
 0
```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

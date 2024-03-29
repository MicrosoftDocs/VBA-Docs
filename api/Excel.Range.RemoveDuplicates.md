---
title: Range.RemoveDuplicates method (Excel)
keywords: vbaxl10.chm144243
f1_keywords:
- vbaxl10.chm144243
api_name:
- Excel.Range.RemoveDuplicates
ms.assetid: 0e74bde2-08b3-898d-0b30-53de911bd7e9
ms.date: 05/11/2019
ms.localizationpriority: medium
---


# Range.RemoveDuplicates method (Excel)

Removes duplicate values from a range of values.

## Syntax

_expression_.**RemoveDuplicates** (_Columns_ , _Header_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Columns_|Required| **Variant**|Array of indexes of the columns that contain the duplicate information. |
| _Header_|Optional| **[XlYesNoGuess](excel.xlyesnoguess.md)**|Specifies whether the first row contains header information. **xlNo** is the default value; specify **xlGuess** if you want Excel to attempt to determine the header.|


## Example

The following code sample removes duplicates with the first 2 columns.

```vb
ActiveSheet.Range("A1:C100").RemoveDuplicates Columns:=Array(1,2), Header:=xlYes
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

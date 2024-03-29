---
title: Range.Consolidate method (Excel)
keywords: vbaxl10.chm144103
f1_keywords:
- vbaxl10.chm144103
api_name:
- Excel.Range.Consolidate
ms.assetid: d5fb78a3-c3ec-0d1a-c6ad-b33bc90e431c
ms.date: 05/10/2019
ms.localizationpriority: medium
---


# Range.Consolidate method (Excel)

Consolidates data from multiple ranges on multiple worksheets into a single range on a single worksheet. **Variant**.


## Syntax

_expression_.**Consolidate** (_Sources_, _Function_, _TopRow_, _LeftColumn_, _CreateLinks_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sources_|Optional| **Variant**|The sources of the consolidation as an array of text reference strings in R1C1-style notation. The references must include the full path of sheets to be consolidated.|
| _Function_|Optional| **Variant**|One of the constants of **[XlConsolidationFunction](Excel.XlConsolidationFunction.md)**, which specifies the type of consolidation.|
| _TopRow_|Optional| **Variant**| **True** to consolidate data based on column titles in the top row of the consolidation ranges. **False** to consolidate data by position. The default value is **False**.|
| _LeftColumn_|Optional| **Variant**| **True** to consolidate data based on row titles in the left column of the consolidation ranges. **False** to consolidate data by position. The default value is **False**.|
| _CreateLinks_|Optional| **Variant**| **True** to have the consolidation use worksheet links. **False** to have the consolidation copy the data. The default value is **False**.|

## Return value

Variant


## Example

This example consolidates data from Sheet2 and Sheet3 onto Sheet1 by using the SUM function.

```vb
Worksheets("Sheet1").Range("A1").Consolidate _ 
 Sources:=Array("Sheet2!R1C1:R37C6", "Sheet3!R1C1:R37C6"), _ 
 Function:=xlSum
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

---
title: VPageBreaks.Add method (Excel)
keywords: vbaxl10.chm168076
f1_keywords:
- vbaxl10.chm168076
ms.prod: excel
api_name:
- Excel.VPageBreaks.Add
ms.assetid: 3196719d-c423-675b-6465-8ac0e9a1c302
ms.date: 05/18/2019
localization_priority: Normal
---


# VPageBreaks.Add method (Excel)

Adds a vertical page break.


## Syntax

_expression_.**Add** (_Before_)

_expression_ A variable that represents a **[VPageBreaks](Excel.VPageBreaks.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Before_|Required| **Object**|A **[Range](Excel.Range(object).md)** object. The range to the left of which the new page break will be added.|

## Return value

A **[VPageBreak](Excel.VPageBreak.md)** object that represents the new vertical page break.


## Example

This example adds a horizontal page break above cell F25 and adds a vertical page break to the left of this cell.

```vb
With Worksheets(1) 
 .HPageBreaks.Add .Range("F25") 
 .VPageBreaks.Add .Range("F25") 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
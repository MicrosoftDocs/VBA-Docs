---
title: PivotFormulas.Add method (Excel)
keywords: vbaxl10.chm233078
f1_keywords:
- vbaxl10.chm233078
ms.prod: excel
api_name:
- Excel.PivotFormulas.Add
ms.assetid: 53969cea-74e5-7102-9a80-89b854006edd
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotFormulas.Add method (Excel)

Creates a new PivotTable formula. 


## Syntax

_expression_.**Add** (_Formula_, _UseStandardFormula_)

_expression_ A variable that represents a **[PivotFormulas](Excel.PivotFormulas.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Formula_|Required| **String**|The new PivotTable formula.|
| _UseStandardFormula_|Optional| **Variant**|A standard PivotTable formula.|

## Return value

A **[PivotFormula](Excel.PivotFormula.md)** object that represents the new PivotTable formula.


## Example

This example creates a new PivotTable formula for the first PivotTable report on worksheet one.

```vb
Worksheets(1).PivotTables(1).PivotFormulas _ 
 .Add "Year['1998'] Apples = (Year['1997'] Apples) * 2"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
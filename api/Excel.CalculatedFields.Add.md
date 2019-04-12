---
title: CalculatedFields.Add method (Excel)
keywords: vbaxl10.chm244078
f1_keywords:
- vbaxl10.chm244078
ms.prod: excel
api_name:
- Excel.CalculatedFields.Add
ms.assetid: 7c01ebbf-d6a4-2b4d-4740-5cb4e2de826a
ms.date: 04/13/2019
localization_priority: Normal
---


# CalculatedFields.Add method (Excel)

Creates a new calculated field. Returns a **[PivotField](Excel.PivotField.md)** object.


## Syntax

_expression_.**Add** (_Name_, _Formula_, _UseStandardFormula_)

_expression_ A variable that represents a **[CalculatedFields](Excel.CalculatedFields.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the field.|
| _Formula_|Required| **String**|The formula for the field.|
| _UseStandardFormula_|Optional| **Variant**| **False** (default) for upward compatibility. **True** for strings contained in any arguments that are field names; will be interpreted as having been formatted in standard U.S. English instead of local settings.|

## Return value

A **PivotField** that represents the new calculated field.


## Example

This example adds a calculated field to the first PivotTable report on worksheet one.

```vb
Worksheets(1).PivotTables(1).CalculatedFields.Add "PxS", _ 
 "= Product * Sales"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

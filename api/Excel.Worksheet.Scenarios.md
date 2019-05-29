---
title: Worksheet.Scenarios method (Excel)
keywords: vbaxl10.chm175123
f1_keywords:
- vbaxl10.chm175123
ms.prod: excel
api_name:
- Excel.Worksheet.Scenarios
ms.assetid: 52e60b55-9316-4c0b-4cb7-ef4605bd31eb
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.Scenarios method (Excel)

Returns an object that represents either a single scenario (a **[Scenario](Excel.Scenario.md)** object) or a collection of scenarios (a **[Scenarios](Excel.Scenarios.md)** object) on the worksheet.


## Syntax

_expression_.**Scenarios** (_Index_)

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the scenario. Use an array to specify more than one scenario.|

## Return value

**Object**


## Example

This example sets the comment for the first scenario on Sheet1.

```vb
Worksheets("Sheet1").Scenarios(1).Comment = _ 
 "Worst-case July 1993 sales"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
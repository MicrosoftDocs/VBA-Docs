---
title: Scenario.ChangingCells property (Excel)
keywords: vbaxl10.chm364074
f1_keywords:
- vbaxl10.chm364074
ms.prod: excel
api_name:
- Excel.Scenario.ChangingCells
ms.assetid: 254abee5-0b64-7f68-33e9-28228541ad8f
ms.date: 05/11/2019
localization_priority: Normal
---


# Scenario.ChangingCells property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents the changing cells for a scenario. Read-only.


## Syntax

_expression_.**ChangingCells**

_expression_ A variable that represents a **[Scenario](Excel.Scenario.md)** object.


## Example

This example selects the changing cells for scenario one on Sheet1.

```vb
Worksheets("Sheet1").Activate 
ActiveSheet.Scenarios(1).ChangingCells.Select
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: Scenario.Comment property (Excel)
keywords: vbaxl10.chm364075
f1_keywords:
- vbaxl10.chm364075
ms.prod: excel
api_name:
- Excel.Scenario.Comment
ms.assetid: 0fe0a22d-b9d0-4e7c-e5db-258a676f222e
ms.date: 05/11/2019
localization_priority: Normal
---


# Scenario.Comment property (Excel)

Returns or sets a **String** value that represents the comment associated with the scenario.


## Syntax

_expression_.**Comment**

_expression_ A variable that represents a **[Scenario](Excel.Scenario.md)** object.


## Remarks

The comment text cannot exceed 255 characters.


## Example

This example sets the comment for scenario one on Sheet1.

```vb
Worksheets("Sheet1").Scenarios(1).Comment = _ 
 "Worst case July 1993 sales"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
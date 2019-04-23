---
title: Scenario object (Excel)
keywords: vbaxl10.chm363072
f1_keywords:
- vbaxl10.chm363072
ms.prod: excel
api_name:
- Excel.Scenario
ms.assetid: edd1c4f4-12b1-0d9f-f4aa-dd66278ba891
ms.date: 04/02/2019
localization_priority: Normal
---


# Scenario object (Excel)

Represents a scenario on a worksheet.


## Remarks

A scenario is a group of input values (called _changing cells_) that's named and saved. The **Scenario** object is a member of the **[Scenarios](Excel.Scenarios.md)** collection. The **Scenarios** collection contains all the defined scenarios for a worksheet.


## Example

Use **[Scenarios](Excel.Worksheet.Scenarios.md)** (_index_), where _index_ is the scenario name or index number, to return a single **Scenario** object. 

The following example shows the scenario named Typical on the worksheet named Options.

```vb
Worksheets("options").Scenarios("typical").Show
```

## Methods

- [ChangeScenario](Excel.Scenario.ChangeScenario.md)
- [Delete](Excel.Scenario.Delete.md)
- [Show](Excel.Scenario.Show.md)

## Properties

- [Application](Excel.Scenario.Application.md)
- [ChangingCells](Excel.Scenario.ChangingCells.md)
- [Comment](Excel.Scenario.Comment.md)
- [Creator](Excel.Scenario.Creator.md)
- [Hidden](Excel.Scenario.Hidden.md)
- [Index](Excel.Scenario.Index.md)
- [Locked](Excel.Scenario.Locked.md)
- [Name](Excel.Scenario.Name.md)
- [Parent](Excel.Scenario.Parent.md)
- [Values](Excel.Scenario.Values.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
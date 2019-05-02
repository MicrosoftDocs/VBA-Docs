---
title: Research object (Excel)
keywords: vbaxl10.chm848072
f1_keywords:
- vbaxl10.chm848072
ms.prod: excel
api_name:
- Excel.Research
ms.assetid: de9d8a1d-4942-88f4-ba8c-30bd06e1f24b
ms.date: 04/02/2019
localization_priority: Normal
---


# Research object (Excel)

Represents the controls of a **Research** query.


## Remarks

When working with **Research** queries, you must have an existing GUID that corresponds to a live data source. If the data source is unavailable or does not exist, a run-time error occurs.


## Example

The following example returns data from an existing data source and translates the information into working content.

```vb
Worksheets("Sheet1").Research.Translate = True
```

## Methods

- [IsResearchService](Excel.Research.IsResearchService.md)
- [Query](Excel.Research.Query.md)
- [SetLanguagePair](Excel.Research.SetLanguagePair.md)

## Properties

- [Application](Excel.Research.Application.md)
- [Creator](Excel.Research.Creator.md)
- [Parent](Excel.Research.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
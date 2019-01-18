---
title: Outline object (Excel)
keywords: vbaxl10.chm454072
f1_keywords:
- vbaxl10.chm454072
ms.prod: excel
api_name:
- Excel.Outline
ms.assetid: f5d50a8a-0dd9-638a-4374-5c648386a598
ms.date: 06/08/2017
localization_priority: Normal
---


# Outline object (Excel)

Represents an outline on a worksheet.


## Example

Use the  **[Outline](Excel.Worksheet.Outline.md)** property to return an **Outline** object. The following example sets the outline on Sheet4 so that only the first outline level is shown.


```vb
Worksheets("sheet4").Outline.ShowLevels 1
```


## Methods



|Name|
|:-----|
|[ShowLevels](Excel.Outline.ShowLevels.md)|

## Properties



|Name|
|:-----|
|[Application](Excel.Outline.Application.md)|
|[AutomaticStyles](Excel.Outline.AutomaticStyles.md)|
|[Creator](Excel.Outline.Creator.md)|
|[Parent](Excel.Outline.Parent.md)|
|[SummaryColumn](Excel.Outline.SummaryColumn.md)|
|[SummaryRow](Excel.Outline.SummaryRow.md)|

## See also


[Excel Object Model Reference](overview/Excel/object-model.md)

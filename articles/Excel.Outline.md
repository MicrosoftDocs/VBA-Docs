---
title: Outline Object (Excel)
keywords: vbaxl10.chm454072
f1_keywords:
- vbaxl10.chm454072
ms.prod: excel
api_name:
- Excel.Outline
ms.assetid: f5d50a8a-0dd9-638a-4374-5c648386a598
ms.date: 06/08/2017
---


# Outline Object (Excel)

Represents an outline on a worksheet.


## Example

Use the  **[Outline](Excel.Worksheet.Outline.md)** property to return an **Outline** object. The following example sets the outline on Sheet4 so that only the first outline level is shown.


```
Worksheets("sheet4").Outline.ShowLevels 1
```


## Methods



|**Name**|
|:-----|
|[ShowLevels](Excel.Outline.ShowLevels.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.Outline.Application.md)|
|[AutomaticStyles](Excel.Outline.AutomaticStyles.md)|
|[Creator](Excel.Outline.Creator.md)|
|[Parent](Excel.Outline.Parent.md)|
|[SummaryColumn](Excel.Outline.SummaryColumn.md)|
|[SummaryRow](Excel.Outline.SummaryRow.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

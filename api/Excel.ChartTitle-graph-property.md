---
title: ChartTitle Property (Excel Graph)
keywords: vbagr10.chm5207199
f1_keywords:
- vbagr10.chm5207199
ms.prod: excel
api_name:
- Excel.ChartTitle
ms.assetid: 736a91ad-a2ef-82c4-33b7-85c5ff78ae08
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartTitle Property (Excel Graph)

Returns a  **[ChartTitle](Excel.ChartTitle-graph-object.md)** object that represents the title of the specified chart. Read-only.


## Example

This example sets the text for the title of the chart.


```vb
With myChart 
 .HasTitle = True 
 .ChartTitle.Text = "First Quarter Sales" 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
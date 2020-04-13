---
title: ChartTitle object (PowerPoint)
keywords: vbapp10.chm694000
f1_keywords:
- vbapp10.chm694000
ms.prod: powerpoint
api_name:
- PowerPoint.ChartTitle
ms.assetid: 21305a3b-1c77-d420-2156-79083189df03
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartTitle object (PowerPoint)

Represents the chart title.


## Remarks

Use the  **[ChartTitle](PowerPoint.Chart.ChartTitle.md)** property to return the **ChartTitle** object.

The **ChartTitle** object does not exist and cannot be used unless the **[HasTitle](PowerPoint.Chart.HasTitle.md)** property for the chart is **True**.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

 The following example adds a title to the first embedded chart in the document.




```vb
With ActiveDocument.InlineShapes(1).Chart

    .HasTitle = True

    .ChartTitle.Text = "February Sales"

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
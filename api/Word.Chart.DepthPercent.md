---
title: Chart.DepthPercent property (Word)
keywords: vbawd10.chm79364100
f1_keywords:
- vbawd10.chm79364100
ms.prod: word
api_name:
- Word.Chart.DepthPercent
ms.assetid: fd1a83dc-e68d-82be-d2bf-5f7a87cb08ac
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.DepthPercent property (Word)

Returns or sets the depth of a 3D chart as a percentage of the chart width (between 20 and 2000 percent). Read/write  **Long**.


## Syntax

_expression_.**DepthPercent**

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Remarks

This property applies only to 3D charts.


## Example

The following example sets the depth of the first chart in the active document to be 50 percent of its width. You should run this example on a 3D chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 Chart.DepthPercent = 50 
 End If 
End With 

```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
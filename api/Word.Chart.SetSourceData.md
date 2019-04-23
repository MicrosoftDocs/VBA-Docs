---
title: Chart.SetSourceData method (Word)
keywords: vbawd10.chm79365509
f1_keywords:
- vbawd10.chm79365509
ms.prod: word
api_name:
- Word.Chart.SetSourceData
ms.assetid: 8c5b056a-6680-7e4e-ce67-a3b76b2d7d25
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.SetSourceData method (Word)

Sets the source data range for the chart.


## Syntax

_expression_.**SetSourceData** (_Source_, _PlotBy_)

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Source_|Required| **String**|The address of the chart data range that contains the source data.|
| _PlotBy_|Optional| **Variant**|Specifies the way the data will be plotted. Can be either of the following  **[XlRowCol](Word.xlrowcol.md)** constants: **xlColumns** or **xlRows**.|

## Example

The following example sets the source data range for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SetSourceData _ 
 Source:="='Sheet1'!$A$1:$D$5", _ 
 PlotBy:=xlColumns 
 End If 
End With
```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
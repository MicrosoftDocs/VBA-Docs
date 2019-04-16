---
title: Chart.HasLegend property (Word)
ms.prod: word
api_name:
- Word.Chart.HasLegend
ms.assetid: 057fedc3-4f23-9c28-3196-836523d83656
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.HasLegend property (Word)

 **True** if the chart has a legend. Read/write **Boolean**.


## Syntax

_expression_.**HasLegend**

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Example

The following example enables the legend for the first chart in the active document and then sets the legend font color to blue.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 .HasLegend = True 
 .Legend.Font.ColorIndex = 5 
 End With 
 End If 
End With
```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: ChartGroup.SizeRepresents property (Word)
keywords: vbawd10.chm263454754
f1_keywords:
- vbawd10.chm263454754
ms.prod: word
api_name:
- Word.ChartGroup.SizeRepresents
ms.assetid: 9611e92a-725c-fbe8-41bf-ef57d2166e4d
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.SizeRepresents property (Word)

Returns or sets what the bubble size represents on a bubble chart. Read/write  **Long**.


## Syntax

_expression_.**SizeRepresents**

_expression_ A variable that represents a **[ChartGroup](Word.ChartGroup.md)** object.


## Remarks

This property can be either of the following  **[XlSizeRepresents](Word.xlsizerepresents.md)** constants:


-  **xlSizeIsArea**
    
-  **xlSizeIsWidth**
    



## Example

The following example sets what the bubble size represents for chart group one of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartGroups(1).SizeRepresents = xlSizeIsWidth 
 End If 
End With
```


## See also


[ChartGroup Object](Word.ChartGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
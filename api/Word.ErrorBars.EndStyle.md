---
title: ErrorBars.EndStyle property (Word)
keywords: vbawd10.chm74843236
f1_keywords:
- vbawd10.chm74843236
ms.prod: word
api_name:
- Word.ErrorBars.EndStyle
ms.assetid: e0396671-1c83-c844-2ec3-e205ffda6ddf
ms.date: 06/08/2017
localization_priority: Normal
---


# ErrorBars.EndStyle property (Word)

Returns or sets the end style for the error bars. Read/write  **Long**.


## Syntax

_expression_.**EndStyle**

_expression_ A variable that represents an '[ErrorBars](Word.ErrorBars.md)' object.


## Remarks

The value of this property can be one of the following  **[XlEndStyleCap](Word.xlendstylecap.md)** constants:


-  **xlCap**
    
-  **xlNoCap**
    



## Example

The following example sets the end style for the error bars for series one of the first chart in the active document. You should run the example on a 2D line chart that has Y error bars for the first series.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).ErrorBars.EndStyle = xlCap 
 End If 
End With
```


## See also


[ErrorBars Object](Word.ErrorBars.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
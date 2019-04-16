---
title: Chart.PlotVisibleOnly property (Word)
ms.prod: word
api_name:
- Word.Chart.PlotVisibleOnly
ms.assetid: 59b7f58e-a1b2-56cd-89e8-529228d2979c
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.PlotVisibleOnly property (Word)

 **True** if only visible cells are plotted. **False** if both visible and hidden cells are plotted. Read/write **Boolean**.


## Syntax

_expression_.**PlotVisibleOnly**

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Example

The following example causes Microsoft Word to plot only visible cells for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.PlotVisibleOnly = True 
 End If 
End With
```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
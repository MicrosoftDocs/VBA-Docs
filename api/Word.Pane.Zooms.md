---
title: Pane.Zooms property (Word)
keywords: vbawd10.chm157286407
f1_keywords:
- vbawd10.chm157286407
ms.prod: word
api_name:
- Word.Pane.Zooms
ms.assetid: 6a09981c-cc68-2468-f750-18cb8524767c
ms.date: 06/08/2017
localization_priority: Normal
---


# Pane.Zooms property (Word)

Returns a  **[Zooms](Word.zooms.md)** collection that represents the magnification options for each view (such as normal view, outline view or print layout view).


## Syntax

_expression_. `Zooms`

 _expression_ An expression that returns a '[Pane](Word.Pane.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example sets the magnification in normal view to 100 percent for each open window.


```vb
Dim wndLoop as Window 
 
For Each wndLoop In Windows 
 wndLoop.ActivePane.Zooms(wdNormalView).Percentage = 100 
Next wndLoop
```

This example sets the magnification in print layout view so that an entire page is visible.




```vb
ActiveDocument.ActiveWindow.Panes(1).Zooms(wdPrintView).PageFit = _ 
 wdPageFitFullPage
```


## See also


[Pane Object](Word.Pane.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
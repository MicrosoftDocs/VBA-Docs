---
title: View.Zoom property (Word)
keywords: vbawd10.chm161808394
f1_keywords:
- vbawd10.chm161808394
ms.prod: word
api_name:
- Word.View.Zoom
ms.assetid: 587c2f80-461a-76f8-35b8-a14f73fb80ef
ms.date: 06/08/2017
localization_priority: Normal
---


# View.Zoom property (Word)

Returns a  **[Zoom](Word.Zoom.md)** object that represents the magnification for the specified view.


## Syntax

 _expression_. `Zoom`

 _expression_ An expression that returns one of a '[View](Word.View.md)' object.


## Example

This example changes the zoom percentage of each open window to 125 percent.


```vb
Sub wndBig() 
 Dim wndBig As Window 
 
 For Each wndBig In Windows 
 wndBig.View.Zoom.Percentage = 125 
 Next wndBig 
End Sub
```

This example changes the zoom percentage of the active window so that the entire width of the text is visible.




```vb
ActiveDocument.ActiveWindow.View.Zoom.PageFit = wdPageFitBestFit
```


## See also


[View Object](Word.View.md)


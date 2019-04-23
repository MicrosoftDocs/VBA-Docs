---
title: Global.PrintPreview property (Word)
keywords: vbawd10.chm163119131
f1_keywords:
- vbawd10.chm163119131
ms.prod: word
api_name:
- Word.Global.PrintPreview
ms.assetid: f9da7e12-0d4b-4d1c-fd53-219f0f9c146f
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.PrintPreview property (Word)

 **True** if print preview is the current view. Read/write **Boolean**.


## Syntax

_expression_. `PrintPreview`

_expression_ Required. A variable that represents a '[Global](Word.Global.md)' object.


## Example

This example switches the view to print preview.


```vb
PrintPreview = True
```

This example switches the active window from print preview to normal view.




```vb
PrintPreview = False 
ActiveDocument.ActiveWindow.View.Type = wdNormalView
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
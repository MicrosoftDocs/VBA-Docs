---
title: System.HorizontalResolution property (Word)
keywords: vbawd10.chm154468359
f1_keywords:
- vbawd10.chm154468359
ms.prod: word
api_name:
- Word.System.HorizontalResolution
ms.assetid: 1e26725e-4914-b9ac-be2d-05991f4c144f
ms.date: 06/08/2017
localization_priority: Normal
---


# System.HorizontalResolution property (Word)

Returns the horizontal display resolution, in pixels. Read-only  **Long**.


## Syntax

 _expression_. `HorizontalResolution`

 _expression_ A variable that represents a '[System](Word.System.md)' object.


## Example

This example displays the current screen resolution (for example, "1024 x 768").


```vb
Dim lngHorizontal As Long 
Dim lngVertical As Long 
 
lngHorizontal = System.HorizontalResolution 
lngVertical = System.VerticalResolution 
MsgBox "Resolution = " & lngHorizontal & " x " & lngVertical
```


## See also


[System Object](Word.System.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
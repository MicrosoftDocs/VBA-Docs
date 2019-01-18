---
title: Border.ArtWidth property (Word)
keywords: vbawd10.chm154861574
f1_keywords:
- vbawd10.chm154861574
ms.prod: word
api_name:
- Word.Border.ArtWidth
ms.assetid: c99ad844-3a47-6291-b84f-d11db78c1f8d
ms.date: 06/08/2017
localization_priority: Normal
---


# Border.ArtWidth property (Word)

Returns or sets the width (in points) of the specified graphical page border. Read/write  **Long**.


## Syntax

 _expression_. `ArtWidth`

 _expression_ A variable that represents a '[Border](Word.Border.md)' object.


## Example

This example adds a 6-point dotted border around each page in the first section in the selection.


```vb
Dim borderLoop As Border 
 
For Each borderLoop In Selection.Sections(1).Borders 
 With borderLoop 
 .ArtStyle = wdArtBasicBlackDots 
 .ArtWidth = 6 
 End With 
Next borderLoop
```


## See also


[Border Object](Word.Border.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: Borders.EnableFirstPageInSection property (Word)
keywords: vbawd10.chm154927134
f1_keywords:
- vbawd10.chm154927134
ms.prod: word
api_name:
- Word.Borders.EnableFirstPageInSection
ms.assetid: 89eae9eb-25ef-f068-1098-f00389b10a22
ms.date: 06/08/2017
localization_priority: Normal
---


# Borders.EnableFirstPageInSection property (Word)

 **True** if page borders are enabled for the first page in the section. Read/write **Boolean**.


## Syntax

_expression_. `EnableFirstPageInSection`

_expression_ A variable that represents a '[Borders](Word.borders.md)' object.


## Example

This example adds a border around the first page in the first section in the selection.


```vb
Dim borderLoop As Border 
 
With Selection.Sections(1) 
 .Borders.EnableFirstPageInSection = True 
 .Borders.EnableOtherPagesInSection = False 
 For Each borderLoop In .Borders 
 borderLoop.ArtStyle = wdArtPeople 
 borderLoop.ArtWidth = 15 
 Next borderLoop 
End With
```


## See also


[Borders Collection Object](Word.borders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
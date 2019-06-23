---
title: Application.CaptionLabels property (Word)
keywords: vbawd10.chm158334996
f1_keywords:
- vbawd10.chm158334996
ms.prod: word
api_name:
- Word.Application.CaptionLabels
ms.assetid: cf59346d-2ff5-938b-52ea-e2931422fd88
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CaptionLabels property (Word)

Returns a  **[CaptionLabels](Word.captionlabels.md)** collection that represents all the available caption labels. Read-only.


## Syntax

_expression_. `CaptionLabels`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example sets the numbering style for table captions.


```vb
CaptionLabels(wdCaptionTable).NumberStyle = _ 
 wdCaptionNumberStyleLowercaseRoman
```

This example adds a new caption label named "Photo" and then inserts a photo caption.




```vb
CaptionLabels.Add Name:="Photo" 
With Selection 
 .InsertParagraphAfter 
 .InsertCaption Label:="Photo" 
End With
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
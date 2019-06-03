---
title: Global.CaptionLabels property (Word)
keywords: vbawd10.chm163119124
f1_keywords:
- vbawd10.chm163119124
ms.prod: word
api_name:
- Word.Global.CaptionLabels
ms.assetid: 619ae4eb-56fb-ec1d-d2b2-4962e6e4fa5e
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.CaptionLabels property (Word)

Returns a  **[CaptionLabels](Word.captionlabels.md)** collection that represents all the available caption labels. Read-only.


## Syntax

_expression_. `CaptionLabels`

_expression_ A variable that represents a '[Global](Word.Global.md)' object.


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


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
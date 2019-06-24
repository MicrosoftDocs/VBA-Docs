---
title: Application.AutoCaptions property (Word)
keywords: vbawd10.chm158334997
f1_keywords:
- vbawd10.chm158334997
ms.prod: word
api_name:
- Word.Application.AutoCaptions
ms.assetid: 6dd68657-3880-76eb-0dc4-91eb58fb0815
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.AutoCaptions property (Word)

Returns an  **[AutoCaptions](Word.autocaptions.md)** collection that represents the captions that are automatically added when items such as tables and pictures are inserted into a document. Read-only.


## Syntax

_expression_. `AutoCaptions`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the name of each item that automatically gets a caption when inserted into the document.


```vb
Dim captionLoop as AutoCaption 
 
For Each captionLoop In AutoCaptions 
 If captionLoop.AutoInsert Then MsgBox captionLoop.Name 
Next captionLoop
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
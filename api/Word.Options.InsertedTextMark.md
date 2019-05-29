---
title: Options.InsertedTextMark property (Word)
keywords: vbawd10.chm162988089
f1_keywords:
- vbawd10.chm162988089
ms.prod: word
api_name:
- Word.Options.InsertedTextMark
ms.assetid: 6c17aa01-2dcb-cf0e-6e8d-bd7f0b254fe8
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.InsertedTextMark property (Word)

Returns or sets how Microsoft Word formats inserted text while change tracking is enabled (the  **TrackRevisions** property is **True**). Read/write **WdInsertedTextMark**.


## Syntax

_expression_. `InsertedTextMark`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Remarks

If change tracking is not enabled, this property is ignored. Use this property with the  **InsertedTextColor** property to control the appearance of inserted text in a document.

The  **ShowRevisions** property must be **True** to see the formatting for inserted text during editing. The **PrintRevisions** property must be **True** in order for Word to use the formatting for inserted text when printing a document.


## Example

This example sets Word to italicize inserted text.


```vb
Options.InsertedTextMark = wdInsertedTextMarkItalic
```

This example sets Word to format inserted text as bold if it isn't already.




```vb
If Options.InsertedTextMark <> wdInsertedTextMarkBold Then 
 Options.InsertedTextMark = wdInsertedTextMarkBold 
Else 
 MsgBox Prompt:="Inserted text is already bold!" 
End If
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
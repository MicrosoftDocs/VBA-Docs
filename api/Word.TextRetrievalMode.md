---
title: TextRetrievalMode object (Word)
keywords: vbawd10.chm2361
f1_keywords:
- vbawd10.chm2361
ms.prod: word
api_name:
- Word.TextRetrievalMode
ms.assetid: b76ad3a6-efc2-4abb-abb4-b8128577bbf2
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRetrievalMode object (Word)

Represents options that control how text is retrieved from a  **Range** object.


## Remarks

Use the **TextRetrievalMode** property to return a **TextRetrievalMode** object. The following example displays the text of the first sentence in the active document, excluding field codes and hidden text.


```vb
With ActiveDocument.Sentences(1).TextRetrievalMode 
 .IncludeHiddenText = False 
 .IncludeFieldCodes = False 
 MsgBox .Parent.Text 
End With
```

Changing the **ViewType**, **IncludeHiddentText**, or **IncludeFieldCodes** property of the **TextRetrievalMode** object doesn't change the screen display. Instead, changing one of these properties determines what text is retrieved from a **Range** object when the **Text** property is used.


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
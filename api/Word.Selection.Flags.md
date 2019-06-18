---
title: Selection.Flags property (Word)
keywords: vbawd10.chm158663058
f1_keywords:
- vbawd10.chm158663058
ms.prod: word
api_name:
- Word.Selection.Flags
ms.assetid: bca92e77-077c-57d0-3012-8c064e93f112
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Flags property (Word)

Returns or sets properties of the selection. Read/write  **WdSelectionFlags**.


## Syntax

_expression_.**Flags**

_expression_ Required. An expression that returns a **[Selection](Word.Selection.md)** object.


## Example

This example selects the first word in the active document. The first message box displays "False" because the end of the selection is active. The  **Flags** property makes the beginning of the selection active, and the second message box displays "True."


```vb
ActiveDocument.Words(1).Select 
MsgBox Selection.StartIsActive 
Selection.Flags = wdSelStartActive 
MsgBox Selection.StartIsActive
```

This example turns on overtype mode for the selection.




```vb
Selection.Flags = wdSelStartActive
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
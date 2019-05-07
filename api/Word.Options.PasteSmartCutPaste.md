---
title: Options.PasteSmartCutPaste property (Word)
keywords: vbawd10.chm162988470
f1_keywords:
- vbawd10.chm162988470
ms.prod: word
api_name:
- Word.Options.PasteSmartCutPaste
ms.assetid: d25143d6-2c83-ce37-3f8e-3177af0eccdd
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.PasteSmartCutPaste property (Word)

 **True** if Microsoft Word intelligently pastes selections into a document. Read/write **Boolean**.


## Syntax

_expression_. `PasteSmartCutPaste`

_expression_ A variable that represents a '[Options](Word.Options.md)' object.


## Example

This example sets Word to enable intelligent selection pasting if the option has been disabled.


```vb
Sub EnableSmartCutPaste() 
 If Options.PasteSmartCutPaste = False Then 
 Options.PasteSmartCutPaste = True 
 End If 
End Sub
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
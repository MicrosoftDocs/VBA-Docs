---
title: Options.SmartParaSelection property (Word)
keywords: vbawd10.chm162988484
f1_keywords:
- vbawd10.chm162988484
ms.prod: word
api_name:
- Word.Options.SmartParaSelection
ms.assetid: 3c3aeb77-febe-b071-03ab-70407ddb58f7
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.SmartParaSelection property (Word)

 **True** for Microsoft Word to include the paragraph mark in a selection when selecting most or all of a paragraph. Read/write **Boolean**.


## Syntax

_expression_. `SmartParaSelection`

_expression_ A variable that represents a '[Options](Word.Options.md)' object.


## Example

This example disables smart paragraph selection.


```vb
Sub SetSmartParagraphSelection() 
 Options.SmartParaSelection = False 
End Sub
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
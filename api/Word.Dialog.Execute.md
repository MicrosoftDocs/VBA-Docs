---
title: Dialog.Execute method (Word)
keywords: vbawd10.chm163085569
f1_keywords:
- vbawd10.chm163085569
ms.prod: word
api_name:
- Word.Dialog.Execute
ms.assetid: 7f7dce3a-40ef-988c-f5ea-06a25c0ccc4b
ms.date: 06/08/2017
localization_priority: Normal
---


# Dialog.Execute method (Word)

Applies the current settings of a Microsoft Word dialog box.


## Syntax

_expression_. `Execute`

_expression_ Required. A variable that represents a '[Dialog](Word.Dialog.md)' object.


## Example

The following example enables the **Keep with next** check box on the **Line and Page Breaks** tab in the **Paragraph** dialog box.


```vb
With Dialogs(wdDialogFormatParagraph) 
 .KeepWithNext = 1 
 .Execute 
End With
```


## See also


[Dialog Object](Word.Dialog.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
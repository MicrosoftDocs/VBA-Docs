---
title: Window.EnvelopeVisible property (Word)
keywords: vbawd10.chm157417505
f1_keywords:
- vbawd10.chm157417505
ms.prod: word
api_name:
- Word.Window.EnvelopeVisible
ms.assetid: d04d6714-ba32-39cc-4853-e9ac6696e718
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.EnvelopeVisible property (Word)

 **True** if the email message header is visible in the document window. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_. `EnvelopeVisible`

_expression_ A variable that represents a **[Window](Word.Window.md)** object.


## Remarks

This property has no effect if the document isn't an email message.


## Example

This example displays the email message header.


```vb
ActiveWindow.EnvelopeVisible = True
```


## See also


[Window Object](Word.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
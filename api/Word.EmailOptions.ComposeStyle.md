---
title: EmailOptions.ComposeStyle property (Word)
keywords: vbawd10.chm165347437
f1_keywords:
- vbawd10.chm165347437
ms.prod: word
api_name:
- Word.EmailOptions.ComposeStyle
ms.assetid: 0c1ada5e-7bf0-2ae1-3223-ed4f76252bb1
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.ComposeStyle property (Word)

Returns a  **[Style](Word.Style.md)** object that represents the style used to compose new email messages. Read-only.


## Syntax

_expression_. `ComposeStyle`

_expression_ A variable that represents a '[EmailOptions](Word.EmailOptions.md)' object.


## Example

This example displays the name of the default style used to compose new email messages.


```vb
MsgBox Application.EmailOptions.ComposeStyle.NameLocal
```

This example changes the font color of the default style used to compose new email messages.




```vb
Application.EmailOptions.ComposeStyle.Font.Color = _ 
 wdColorBrightGreen
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
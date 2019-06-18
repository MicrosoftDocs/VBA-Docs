---
title: Options.CtrlClickHyperlinkToOpen property (Word)
keywords: vbawd10.chm162988467
f1_keywords:
- vbawd10.chm162988467
ms.prod: word
api_name:
- Word.Options.CtrlClickHyperlinkToOpen
ms.assetid: 2180e99c-ab4c-3f75-2417-22cec6b2d130
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.CtrlClickHyperlinkToOpen property (Word)

**True** if Microsoft Word requires holding down the Ctrl key while clicking to open a hyperlink. Read/write **Boolean**.


## Syntax

_expression_. `CtrlClickHyperlinkToOpen`

_expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example disables the option that requires holding down the Ctrl key while clicking hyperlinks to open them.


```vb
Sub ToggleHyperlinkOption() 
 Options.CtrlClickHyperlinkToOpen = False 
End Sub
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
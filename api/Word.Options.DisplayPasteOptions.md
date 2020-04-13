---
title: Options.DisplayPasteOptions property (Word)
keywords: vbawd10.chm162988471
f1_keywords:
- vbawd10.chm162988471
ms.prod: word
api_name:
- Word.Options.DisplayPasteOptions
ms.assetid: 518789bd-4a9e-a3c7-0fab-16e44f63e68d
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.DisplayPasteOptions property (Word)

 **True** for Microsoft Word to display the **Paste Options** button, which displays directly under newly pasted text. Read/write **Boolean**.


## Syntax

_expression_. `DisplayPasteOptions`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example enables the **Paste Options** button if the option has been disabled.


```vb
Sub ShowPasteOptionsButton() 
 With Options 
 If .DisplayPasteOptions = False Then 
 .DisplayPasteOptions = True 
 End If 
 End With 
End Sub
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
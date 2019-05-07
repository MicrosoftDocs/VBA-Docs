---
title: Global.FontNames property (Word)
keywords: vbawd10.chm163119115
f1_keywords:
- vbawd10.chm163119115
ms.prod: word
api_name:
- Word.Global.FontNames
ms.assetid: aa70c33b-2ca3-849a-54b0-fe050072f9ac
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.FontNames property (Word)

Returns a  **[FontNames](Word.FontNames.md)** object that includes the names of all the available fonts. Read-only.


## Syntax

_expression_. `FontNames`

_expression_ A variable that represents a '[Global](Word.Global.md)' object.


## Example

This example displays the font names in the FontNames collection.


```vb
Dim strFont As String 
Dim intResponse As Integer 
 
For Each strFont In FontNames 
 intResponse = MsgBox(Prompt:=strFont, Buttons:=vbOKCancel) 
 If intResponse = vbCancel Then Exit For 
Next strFont
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
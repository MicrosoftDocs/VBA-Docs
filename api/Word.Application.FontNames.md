---
title: Application.FontNames property (Word)
keywords: vbawd10.chm158334987
f1_keywords:
- vbawd10.chm158334987
ms.prod: word
api_name:
- Word.Application.FontNames
ms.assetid: 6aeadf51-79c7-1123-ea64-582ceee26443
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.FontNames property (Word)

Returns a  **[FontNames](Word.FontNames.md)** object that includes the names of all the available fonts. Read-only.


## Syntax

_expression_. `FontNames`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example displays the font names in the **FontNames** collection.


```vb
Dim strFont As String 
Dim intResponse As Integer 
 
For Each strFont In FontNames 
 intResponse = MsgBox(Prompt:=strFont, Buttons:=vbOKCancel) 
 If intResponse = vbCancel Then Exit For 
Next strFont
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
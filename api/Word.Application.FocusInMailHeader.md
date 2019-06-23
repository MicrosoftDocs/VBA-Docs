---
title: Application.FocusInMailHeader property (Word)
keywords: vbawd10.chm158335362
f1_keywords:
- vbawd10.chm158335362
ms.prod: word
api_name:
- Word.Application.FocusInMailHeader
ms.assetid: fba9d08b-1950-b825-5f1a-14d671181b22
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.FocusInMailHeader property (Word)

 **True** if the insertion point is in an email header field (the To: field, for example). Read-only **Boolean**.


## Syntax

_expression_. `FocusInMailHeader`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example displays a message in the status bar if the insertion point is in an email header field.


```vb
If Application.FocusInMailHeader = True Then 
 StatusBar = "Selection is in message header" 
End If
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
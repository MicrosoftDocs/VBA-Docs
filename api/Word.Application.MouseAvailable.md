---
title: Application.MouseAvailable property (Word)
keywords: vbawd10.chm158335013
f1_keywords:
- vbawd10.chm158335013
ms.prod: word
api_name:
- Word.Application.MouseAvailable
ms.assetid: 25ad78ad-c267-35ec-9124-0496c034fa50
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.MouseAvailable property (Word)

 **True** if there is a mouse available for the system. Read-only **Boolean**.


## Syntax

 _expression_. `MouseAvailable`

 _expression_ An expression that returns an '[Application](Word.Application.md)' object.


## Example

This example displays a message that no mouse is available.


```vb
If Application.MouseAvailable = False Then 
 Msgbox "Make sure your mouse is plugged in." 
Else 
 Msgbox "Mouse is available" 
End If
```


## See also


[Application Object](Word.Application.md)


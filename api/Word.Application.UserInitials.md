---
title: Application.UserInitials property (Word)
keywords: vbawd10.chm158335029
f1_keywords:
- vbawd10.chm158335029
ms.prod: word
api_name:
- Word.Application.UserInitials
ms.assetid: 00f7d562-4ce5-00e1-bf9d-4325d47947b2
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.UserInitials property (Word)

Returns or sets the user's initials, which Microsoft Word uses to construct comment marks. Read/write  **String**.


## Syntax

_expression_. `UserInitials`

 _expression_ An expression that returns an **[Application](Word.Application.md)** object. 


## Example

This example sets the user's initials.


```vb
Application.UserInitials = "baa"
```

This example returns the letters found in the  **Initials** box on the **User Information** tab in the **Options** dialog box (**Tools** menu).




```vb
Msgbox Application.UserInitials
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
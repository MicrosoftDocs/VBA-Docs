---
title: Presentation.WritePassword property (PowerPoint)
keywords: vbapp10.chm583081
f1_keywords:
- vbapp10.chm583081
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.WritePassword
ms.assetid: 42381e81-c5d0-3db1-f214-6619bbc6711f
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.WritePassword property (PowerPoint)

Sets or returns the password for saving changes to the specified document. Read/write.


## Syntax

_expression_. `WritePassword`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

String


## Example

This example sets the password for saving changes to the active presentation.


```vb
Sub SetSavePassword()

    ActivePresentation.WritePassword = complexstrPWD 'global variable

End Sub
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
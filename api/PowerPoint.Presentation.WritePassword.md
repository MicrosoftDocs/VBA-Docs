---
title: Presentation.WritePassword property (PowerPoint)
keywords: vbapp10.chm583081
f1_keywords:
- vbapp10.chm583081
api_name:
- PowerPoint.Presentation.WritePassword
ms.assetid: 42381e81-c5d0-3db1-f214-6619bbc6711f
ms.date: 08/02/2022
ms.localizationpriority: medium
---


# Presentation.WritePassword property (PowerPoint)

Sets or returns the password for saving changes to the specified document. Read/write.


## Syntax

_expression_. `WritePassword`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

String


## Remarks

If the presentation is not fully downloaded, the setting of this property fails and an error occurs. For more information about the Partial Documents, see [Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md).


## Example

This example sets the password for saving changes to the active presentation.


```vb
Sub SetSavePassword()

    ActivePresentation.WritePassword = complexstrPWD 'global variable

End Sub
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
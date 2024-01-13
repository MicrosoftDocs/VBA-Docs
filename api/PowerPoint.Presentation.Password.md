---
title: Presentation.Password property (PowerPoint)
keywords: vbapp10.chm583080
f1_keywords:
- vbapp10.chm583080
api_name:
- PowerPoint.Presentation.Password
ms.assetid: 977876b7-b40f-de45-c259-e91744915085
ms.date: 08/02/2022
ms.localizationpriority: medium
---


# Presentation.Password property (PowerPoint)

Returns or sets the password that must be supplied to open the specified presentation. Read/write.


## Syntax

_expression_. `Password`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

String


## Remarks

If the presentation is not fully downloaded, the setting of this property fails and an error occurs. For more information about the Partial Documents, see [Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md).


## Example

This example opens Earnings.ppt, sets a password for it, and then closes the presentation.


```vb
Sub SetPassword()

    With Presentations.Open(FileName:="C:\My Documents\Earnings.ppt")

        .Password = complexstrPWD 'global variable

        .Save

        .Close

    End With

End Sub
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
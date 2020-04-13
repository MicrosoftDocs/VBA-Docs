---
title: TextFrame.Creator property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.Creator
ms.assetid: 7e198a9e-38eb-6f1a-38f6-e24bcac43190
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame.Creator property (PowerPoint)

Returns a **Long** that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[TextFrame](PowerPoint.TextFrame.md)** object.


## Return value

Long


## Remarks

The **Creator** property is designed to be used in Microsoft Office applications for the Macintosh.


## Example

This example displays a message about the creator of myObject.


```vb
Set myObject = Application.ActivePresentation.Slides(1).Shapes(1)

If myObject.Creator = &h50575054 Then

    MsgBox "This is a PowerPoint object"

Else

    MsgBox "This is not a PowerPoint object"

End If
```


## See also


[TextFrame Object](PowerPoint.TextFrame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
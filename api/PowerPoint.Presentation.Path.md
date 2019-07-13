---
title: Presentation.Path property (PowerPoint)
keywords: vbapp10.chm583026
f1_keywords:
- vbapp10.chm583026
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Path
ms.assetid: 67611b54-bc31-ec2b-e645-cb3d4195bbe9
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.Path property (PowerPoint)

Returns a  **String** that represents the path to the specified **[Presentation](PowerPoint.Presentation.md)** object. Read-only.


## Syntax

_expression_.**Path**

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

String


## Remarks

If you use this property to return a path for a presentation that has not been saved, it returns an empty string.

The path doesn't include the final backslash (\) or the name of the specified object. Use the  **Name** property of the **Presentation** object to return the file name without the path, and use the **FullName** property to return the file name and the path together.


## Example

This example saves the active presentation in the same folder as PowerPoint. 


```vb
With Application

    fName = .Path & "\test presentation"

    ActivePresentation.SaveAs fName

End With
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
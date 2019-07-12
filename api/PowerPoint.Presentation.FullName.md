---
title: Presentation.FullName property (PowerPoint)
keywords: vbapp10.chm583024
f1_keywords:
- vbapp10.chm583024
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.FullName
ms.assetid: cf6c5687-5dd0-3e71-3aa9-a370534c4117
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.FullName property (PowerPoint)

Returns the name of the specified add-in or saved presentation, including the path, the current file system separator, and the file name extension. Read-only  **String**.


## Syntax

_expression_.**FullName**

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

String


## Remarks

This property is equivalent to the  **Path** property, followed by the current file system separator, followed by the **Name** property.


## Example

This example displays the path and file name of every available add-in.


```vb
For Each a In Application.AddIns

    MsgBox a.FullName

Next a
```

This example displays the path and file name of the active presentation (assuming that the presentation has been saved).




```vb
MsgBox Application.ActivePresentation.FullName
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
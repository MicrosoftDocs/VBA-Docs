---
title: Presentation.RemovePersonalInformation property (PowerPoint)
keywords: vbapp10.chm583068
f1_keywords:
- vbapp10.chm583068
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.RemovePersonalInformation
ms.assetid: beb422cc-23c5-5de5-ed6f-0fc71315daec
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.RemovePersonalInformation property (PowerPoint)

Determines whether Microsoft PowerPoint should remove all user information from comments and revisions upon saving a presentation. Read/write.


## Syntax

_expression_. `RemovePersonalInformation`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

MsoTriState


## Remarks

The value of the  **RemovePersonalInformation** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**| Comments, revisions, and personal information remain in the presentation.|
|**msoTrue**| Removes comments, revisions, and personal information when saving presentation.|

## Example

This example sets the active presentation to remove personal information the next time the user saves it.


```vb
Sub RemovePersonalInfo()

    ActivePresentation.RemovePersonalInformation = msoTrue

End Sub
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
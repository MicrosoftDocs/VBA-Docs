---
title: Presentation.CustomDocumentProperties property (PowerPoint)
keywords: vbapp10.chm583021
f1_keywords:
- vbapp10.chm583021
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.CustomDocumentProperties
ms.assetid: 3f972f15-f606-0a11-56b6-1994e617def2
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.CustomDocumentProperties property (PowerPoint)

Returns a **DocumentProperties** collection that represents all the custom document properties for the specified presentation. Read-only.

> [!NOTE] 
> All custom document properties will be lost when user uses the **Design Ideas** to change the look of a slide in the presentation.


## Syntax

_expression_. `CustomDocumentProperties`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

DocumentProperties


## Remarks

Use the **[BuiltInDocumentProperties](PowerPoint.Presentation.BuiltInDocumentProperties.md)** property to return the collection of built-in document properties.

For information about returning a single member of a collection, see [Returning an object from a collection](../powerpoint/How-to/return-objects-from-collections.md).


## Example

This example adds a static custom property named "Complete" for the active presentation.


```vb
Application.ActivePresentation.CustomDocumentProperties _
    .Add Name:="Complete", LinkToContent:=False, _
    Type:=msoPropertyTypeBoolean, Value:=False
```

This example displays the active presentation if the value of the "Complete" custom property is **True**.




```vb
With Application.ActivePresentation

    If .CustomDocumentProperties("complete") Then .PrintOut

End With
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

---
title: Shape.Type property (PowerPoint)
keywords: vbapp10.chm547038
f1_keywords:
- vbapp10.chm547038
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Type
ms.assetid: 3a6aa03d-8d93-9a08-ef42-8f128ada7b87
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Type property (PowerPoint)

Represents the type of shape or shapes in a range of shapes. Read-only.


## Syntax

_expression_.**Type**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

MsoShapeType


## Remarks

The value of the **Type** property can be one of the **[MsoShapeType](office.msoshapetype.md)** constants.


## Example

This example loops through all the shapes on all the slides in the active presentation and sets all linked Microsoft Office Excel worksheets to be updated manually.


```vb
For Each sld In ActivePresentation.Slides 
    For Each sh In sld.Shapes 
        If sh.Type = msoLinkedOLEObject Then 
            If sh.OLEFormat.ProgID = "Excel.Sheet" Then 
                sh.LinkFormat.AutoUpdate = ppUpdateOptionManual 
            End If 
        End If 
    Next 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

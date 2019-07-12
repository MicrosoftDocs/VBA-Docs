---
title: LinkFormat object (PowerPoint)
keywords: vbapp10.chm563000
f1_keywords:
- vbapp10.chm563000
ms.prod: powerpoint
api_name:
- PowerPoint.LinkFormat
ms.assetid: e89ee344-4197-ac0d-dd53-966e4672a3ce
ms.date: 06/08/2017
localization_priority: Normal
---


# LinkFormat object (PowerPoint)

Contains properties and methods that apply to linked OLE objects, linked pictures, and IIRC media objects. 


## Example

Use the  **LinkFormat** property to return a **LinkFormat** object. The following example loops through all the shapes on all the slides in the active presentation and sets all linked Microsoft Excel worksheets to be updated manually.


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


## Methods



|Name|
|:-----|
|[BreakLink](PowerPoint.LinkFormat.BreakLink.md)|
|[Update](PowerPoint.LinkFormat.Update.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.LinkFormat.Application.md)|
|[AutoUpdate](PowerPoint.LinkFormat.AutoUpdate.md)|
|[Parent](PowerPoint.LinkFormat.Parent.md)|
|[SourceFullName](PowerPoint.LinkFormat.SourceFullName.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
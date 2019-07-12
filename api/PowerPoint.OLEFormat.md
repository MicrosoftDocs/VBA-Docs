---
title: OLEFormat object (PowerPoint)
keywords: vbapp10.chm562000
f1_keywords:
- vbapp10.chm562000
ms.prod: powerpoint
api_name:
- PowerPoint.OLEFormat
ms.assetid: fbb6d6dd-4dbb-461b-986e-5095c6dc1486
ms.date: 06/08/2017
localization_priority: Normal
---


# OLEFormat object (PowerPoint)

Contains properties and methods that apply to OLE objects. 


## Remarks

The  **[LinkFormat](PowerPoint.LinkFormat.md)** object contains properties and methods that apply to linked OLE objects only. The **[PictureFormat](PowerPoint.PictureFormat.md)** object contains properties and methods that apply to pictures and OLE objects.


## Example

Use the  **OLEFormat** property to return an **OLEFormat** object. The following example loops through all the shapes on all the slides in the active presentation and sets all linked Microsoft Excel worksheets to be updated manually.


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
|[Activate](PowerPoint.OLEFormat.Activate.md)|
|[DoVerb](PowerPoint.OLEFormat.DoVerb.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.OLEFormat.Application.md)|
|[FollowColors](PowerPoint.OLEFormat.FollowColors.md)|
|[Object](PowerPoint.OLEFormat.Object.md)|
|[ObjectVerbs](PowerPoint.OLEFormat.ObjectVerbs.md)|
|[Parent](PowerPoint.OLEFormat.Parent.md)|
|[ProgID](PowerPoint.OLEFormat.ProgID.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
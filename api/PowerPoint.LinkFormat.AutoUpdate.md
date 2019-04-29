---
title: LinkFormat.AutoUpdate property (PowerPoint)
keywords: vbapp10.chm563004
f1_keywords:
- vbapp10.chm563004
ms.prod: powerpoint
api_name:
- PowerPoint.LinkFormat.AutoUpdate
ms.assetid: de142aa6-2414-61c3-62d1-1226a0f9209f
ms.date: 06/08/2017
localization_priority: Normal
---


# LinkFormat.AutoUpdate property (PowerPoint)

Returns or sets the way the link will be updated. Read/write.


## Syntax

_expression_.**AutoUpdate**

_expression_ A variable that represents a **[LinkFormat](PowerPoint.LinkFormat.md)** object.


## Return value

PpUpdateOption


## Remarks

The value of the  **AutoUpdate** property can be one of these **PpUpdateOption** constants.



|Constant|Description|
|:-----|:-----|
|**ppUpdateOptionAutomatic**|The link is updated each time the presentation is opened or the source file changes.|
|**ppUpdateOptionManual**| The link is updated only when the user specifically asks to update the presentation.|
|**ppUpdateOptionMixed**||

## Example

This example loops through all the shapes on all the slides in the active presentation and sets all linked Microsoft Excel worksheets to be updated manually.


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


## See also


[LinkFormat Object](PowerPoint.LinkFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
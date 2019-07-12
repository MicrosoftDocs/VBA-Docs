---
title: CalloutFormat.PresetDrop method (PowerPoint)
keywords: vbapp10.chm559005
f1_keywords:
- vbapp10.chm559005
ms.prod: powerpoint
api_name:
- PowerPoint.CalloutFormat.PresetDrop
ms.assetid: e0f99665-4619-334a-a7bb-e53d5f8ef5ec
ms.date: 06/08/2017
localization_priority: Normal
---


# CalloutFormat.PresetDrop method (PowerPoint)

Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that's a specified distance from the top or bottom of the text box.


## Syntax

_expression_.**PresetDrop** (_DropType_)

_expression_ A variable that represents a [CalloutFormat](PowerPoint.CalloutFormat.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DropType_|Required|**MsoCalloutDropType**|The starting position of the callout line relative to the text bounding box.|

## Remarks

The  _DropType_ parameter value can be one of the following **MsoCalloutDropType** constants. Passing **msoCalloutDropCustom** will cause your code to fail.


||
|:-----|
|**msoCalloutDropBottom**|
|**msoCalloutDropCenter**|
|**msoCalloutDropCustom**|
|**msoCalloutDropMixed**|
|**msoCalloutDropTop**|

## Example

This example specifies that the callout line attach to the top of the text bounding box for shape one on _myDocument_. For the example to work, shape one must be a callout.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).Callout.PresetDrop msoCalloutDropTop
```

This example switches between two preset drops for shape one on _myDocument_. For the example to work, shape one must be a callout.




```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).Callout

    If .DropType = msoCalloutDropTop Then

        .PresetDrop msoCalloutDropBottom

    ElseIf .DropType = msoCalloutDropBottom Then

        .PresetDrop msoCalloutDropTop

    End If

End With
```


## See also


[CalloutFormat Object](PowerPoint.CalloutFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
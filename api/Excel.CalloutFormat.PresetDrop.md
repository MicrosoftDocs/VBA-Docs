---
title: CalloutFormat.PresetDrop method (Excel)
keywords: vbaxl10.chm104005
f1_keywords:
- vbaxl10.chm104005
ms.prod: excel
api_name:
- Excel.CalloutFormat.PresetDrop
ms.assetid: 48d67cad-d93b-2b69-35dd-c3de70340a42
ms.date: 04/13/2019
localization_priority: Normal
---


# CalloutFormat.PresetDrop method (Excel)

Specifies whether the callout line attaches to the top, bottom, or center of the callout text box, or whether it attaches at a point that's a specified distance from the top or bottom of the text box.


## Syntax

_expression_.**PresetDrop** (_DropType_)

_expression_ A variable that represents a **[CalloutFormat](Excel.CalloutFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DropType_|Required| **[MsoCalloutDropType](Office.MsoCalloutDropType.md)**|The starting position of the callout line relative to the text bounding box.|

## Example

This example specifies that the callout line attach to the top of the text bounding box for shape one on _myDocument_. For the example to work, shape one must be a callout.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).Callout.PresetDrop msoCalloutDropTop
```

<br/>

This example toggles between two preset drops for shape one on _myDocument_. For the example to work, shape one must be a callout.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).Callout 
    If .DropType = msoCalloutDropTop Then 
        .PresetDrop msoCalloutDropBottom 
    ElseIf .DropType = msoCalloutDropBottom Then 
        .PresetDrop msoCalloutDropTop 
    End If 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
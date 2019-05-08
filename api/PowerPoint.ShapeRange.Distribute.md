---
title: ShapeRange.Distribute method (PowerPoint)
keywords: vbapp10.chm548064
f1_keywords:
- vbapp10.chm548064
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Distribute
ms.assetid: bbabe9db-30ba-e165-0dcc-7a15e849228e
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.Distribute method (PowerPoint)

Evenly distributes the shapes in the specified range of shapes. You can specify whether you want to distribute the shapes horizontally or vertically and whether you want to distribute them over the entire slide or just over the space they originally occupy.


## Syntax

_expression_. `Distribute`( `_DistributeCmd_`, `_RelativeTo_` )

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DistributeCmd_|Required|**[MsoDistributeCmd](Office.MsoDistributeCmd.md)**|Specifies whether shapes in the range are to be distributed horizontally or vertically.|
| _RelativeTo_|Required|**[MsoTriState](Office.MsoTriState.md)**|Determines whether shapes are distributed evenly over the entire horizontal or vertical space on the slide.|

## Return value

Nothing


## Example

This example defines a shape range that contains all the AutoShapes on the  _myDocument_ and then horizontally distributes the shapes in this range.


```vb
Set myDocument = ActivePresentation.Slides(1) 
With myDocument.Shapes 
    numShapes = .Count 
    If numShapes > 1 Then 
        numAutoShapes = 0 
        ReDim autoShpArray(1 To numShapes) 
        For i = 1 To numShapes 
            If .Item(i).Type = msoAutoShape Then 
                numAutoShapes = numAutoShapes + 1 
                autoShpArray(numAutoShapes) = .Item(i).Name 
            End If 
        Next 
        If numAutoShapes > 1 Then 
            ReDim Preserve autoShpArray(1 To numAutoShapes) 
            Set asRange = .Range(autoShpArray) 
            asRange.Distribute msoDistributeHorizontally, msoFalse 
        End If 
    End If 
End With
```


## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: Shapes.AddConnector method (Excel)
keywords: vbaxl10.chm638078
f1_keywords:
- vbaxl10.chm638078
ms.prod: excel
api_name:
- Excel.Shapes.AddConnector
ms.assetid: 7ea648eb-ac6b-981d-652b-40cea1b3a8da
ms.date: 05/15/2019
localization_priority: Normal
---


# Shapes.AddConnector method (Excel)

Creates a connector. Returns a **[Shape](Excel.Shape.md)** object that represents the new connector. When a connector is added, it's not connected to anything. Use the **[BeginConnect](Excel.ConnectorFormat.BeginConnect.md)** and **[EndConnect](Excel.ConnectorFormat.EndConnect.md)** methods to attach the beginning and end of a connector to other shapes in the document.


## Syntax

_expression_.**AddConnector** (_Type_, _BeginX_, _BeginY_, _EndX_, _EndY_)

_expression_ A variable that represents a **[Shapes](Excel.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[MsoConnectorType](Office.MsoConnectorType.md)**|The connector type to add.|
| _BeginX_|Required| **Single**|The horizontal position (in [points](../language/glossary/vbe-glossary.md#point)) of the connector's starting point relative to the upper-left corner of the document.|
| _BeginY_|Required| **Single**|The vertical position (in points) of the connector's starting point relative to the upper-left corner of the document.|
| _EndX_|Required| **Single**|The horizontal position (in points) of the connector's end point relative to the upper-left corner of the document.|
| _EndY_|Required| **Single**|The vertical position (in points) of the connector's end point relative to the upper-left corner of the document.|

## Return value

**Shape**


## Remarks

When you attach a connector to a shape, the size and position of the connector are automatically adjusted, if necessary. Therefore, if you are going to attach a connector to other shapes, the position and dimensions that you specify when adding the connector are irrelevant.


## Example

The following example adds a curved connector to a new canvas in a new worksheet.


```vb
Sub AddCanvasConnector() 
 
    Dim wksNew As Worksheet 
    Dim shpCanvas As Shape 
 
    Set wksNew = Worksheets.Add 
 
    'Add drawing canvas to new worksheet 
    Set shpCanvas = wksNew.Shapes.AddCanvas( _ 
        Left:=150, Top:=150, Width:=200, Height:=300) 
 
    'Add connector to the drawing canvas 
    shpCanvas.CanvasItems.AddConnector _ 
        Type:=msoConnectorStraight, BeginX:=150, _ 
        BeginY:=150, EndX:=200, EndY:=200 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

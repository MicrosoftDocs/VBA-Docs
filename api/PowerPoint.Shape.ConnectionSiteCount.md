---
title: Shape.ConnectionSiteCount property (PowerPoint)
keywords: vbapp10.chm547019
f1_keywords:
- vbapp10.chm547019
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.ConnectionSiteCount
ms.assetid: 2180bb96-d205-03f3-1ace-355f34286b2e
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.ConnectionSiteCount property (PowerPoint)

Returns the number of connection sites on the specified shape. Read-only.


## Syntax

_expression_.**ConnectionSiteCount**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

Long


## Example

This example adds two rectangles to _myDocument_ and joins them with two connectors. The beginnings of both connectors attach to connection site one on the first rectangle; the ends of the connectors attach to the first and last connection sites of the second rectangle.


```vb
Set myDocument = ActivePresentation.Slides(1)

Set s = myDocument.Shapes

Set firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)

Set secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)

lastsite = secondRect.ConnectionSiteCount

With s.AddConnector(msoConnectorCurve, 0, 0, 100, 100) _
        .ConnectorFormat

    .BeginConnect ConnectedShape:=firstRect, ConnectionSite:=1
    .EndConnect ConnectedShape:=secondRect, ConnectionSite:=1

End With

With s.AddConnector(msoConnectorCurve, 0, 0, 100, 100) _
        .ConnectorFormat

    .BeginConnect ConnectedShape:=firstRect, ConnectionSite:=1
    .EndConnect ConnectedShape:=secondRect, _
        ConnectionSite:=lastsite

End With
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
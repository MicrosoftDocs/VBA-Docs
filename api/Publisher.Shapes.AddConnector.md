---
title: Shapes.AddConnector method (Publisher)
keywords: vbapb10.chm2162705
f1_keywords:
- vbapb10.chm2162705
ms.prod: publisher
api_name:
- Publisher.Shapes.AddConnector
ms.assetid: fd1ef969-7960-2555-e355-9804c86f6c01
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddConnector method (Publisher)

Adds a new **[Shape](Publisher.Shape.md)** object representing a connector to the specified **Shapes** collection.


## Syntax

_expression_.**AddConnector** (_Type_, _BeginX_, _BeginY_, _EndX_, _EndY_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Type_|Required| **[MsoConnectorType](office.msoconnectortype.md)**|The type of connector to add. Can be one of the **MsoConnectorType** constants, except for **msoConnectorTypeMixed**, which is not used with this method.|
|_BeginX_|Required| **Variant**|The x-coordinate of the beginning point of the connector.|
|_BeginY_|Required| **Variant**|The y-coordinate of the beginning point of the connector.|
|_EndX_|Required| **Variant**|The x-coordinate of the ending point of the connector.|
|_EndY_|Required| **Variant**|The y-coordinate of the ending point of the connector.|

## Return value

Shape


## Remarks

For the _BeginX_, _BeginY_, _EndX_, and _EndY_ parameters, numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

The new connector isn't connected to any other shape; use the **[BeginConnect](Publisher.ConnectorFormat.BeginConnect.md)** and **[EndConnect](Publisher.ConnectorFormat.EndConnect.md)** methods to connect the new connector to another shape.

## Example

The following example adds a new straight-line connector to the first page of the active publication.

```vb
Dim shpConnect As Shape 
 
Set shpConnect = ActiveDocument.Pages(1).Shapes.AddConnector _ 
 (Type:=msoConnectorStraight, _ 
 BeginX:=144, BeginY:=144, _ 
 EndX:=180, EndY:=72)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
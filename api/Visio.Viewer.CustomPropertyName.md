---
title: Viewer.CustomPropertyName property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.CustomPropertyName
ms.assetid: 6cd7838b-9c7b-0f07-e94b-c24dc800b2d2
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.CustomPropertyName property (Visio Viewer)

Gets the name of the shape data item (custom property) at the specified index position for the specified shape in Microsoft Visio Viewer. Read-only.


## Syntax

_expression_.**CustomPropertyName** (_ShapeIndex_, _PropertyIndex_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ShapeIndex_|Required| **Long**|The index of the shape that contains the specified shape data item (custom property).|
|_PropertyIndex_|Required| **Long**|The index of the shape data item (custom property).|

## Return value

**String**


## Remarks

In versions of Visio prior to Microsoft Office Visio 2007, shape data items were called custom properties.


## Example

The following code gets the name of the first shape data item assigned to the first shape in the collection of shapes on the current page in Visio Viewer.

```vb
Dim strPropertyName As String

strPropertyName = vsoViewer.CustomPropertyName(1, 1)

Debug.Print strPropertyName

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
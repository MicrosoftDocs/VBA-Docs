---
title: Viewer.PageColor property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.PageColor
ms.assetid: afda33d5-461b-44d0-a611-df26c632ce12
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.PageColor property (Visio Viewer)

Gets or sets the color of the page in the current drawing that is open in Microsoft Visio Viewer. Read/write.


## Syntax

_expression_.**PageColor**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**OLE_COLOR**


## Remarks

This property returns a value of data type **OLE_COLOR** that represents the color of the page in Visio Viewer. The **OLE_COLOR** data type is used for properties that return colors.

Valid hexadecimal values for an **OLE_COLOR** data type in Visio Viewer are of the form _&Hbbggrr_, where _bb_ is the blue value, _gg_ the green value, and _rr_ the red value. All three color values range between 00 and FF hexadecimal (255 decimal).

The default value of the **PageColor** property is white (16777215).


## Example

The following example sets the page color to red.

```vb
vsoViewer.PageColor = 225
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
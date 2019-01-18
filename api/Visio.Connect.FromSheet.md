---
title: Connect.FromSheet Property (Visio)
keywords: vis_sdr.chm10313590
f1_keywords:
- vis_sdr.chm10313590
ms.prod: visio
api_name:
- Visio.Connect.FromSheet
ms.assetid: 621aa755-3d17-4c3c-118f-7513d3926b52
ms.date: 06/08/2017
localization_priority: Normal
---


# Connect.FromSheet Property (Visio)

Returns the shape from which a connection or connections originate. Read-only.


## Syntax

 _expression_. `FromSheet`

 _expression_ A variable that represents a [Connect](./Visio.Connect.md) object.


## Return value

Shape


## Remarks

The  **FromSheet** property for a **Connect** object is straightforward. It always returns the shape from which the **Connect** object originates.

A  **Connects** collection represents several connections. If every connection represented by the collection originates from the same shape, the **FromSheet** property for the collection returns that shape. Otherwise, the **FromSheet** property returns **Nothing** and does not raise an exception.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **FromSheet** property to find the shape a **Connect** object originates from in a Microsoft Visio drawing. The example displays the connection information in the Immediate window.

This example assumes there is an active document that contains at least two connected shapes. For best results, connect two shapes from the  **Organization Chart Shapes** stencil.




```vb
 
Public Sub FromSheet_Example() 
 
 Dim vsoShapes As Visio.Shapes 
 Dim vsoShape As Visio.Shape 
 Dim vsoConnectFrom As Visio.Shape 
 Dim vsoConnects As Visio.Connects 
 Dim vsoConnect As Visio.Connect 
 Dim intCurrentShapeIndex As Integer 
 Dim intCounter As Integer 
 
 Set vsoShapes = ActivePage.Shapes 
 
 For intCurrentShapeIndex = 1 To vsoShapes.Count 
 Set vsoShape = vsoShapes(intCurrentShapeIndex) 
 Set vsoConnects = vsoShape.Connects 
 
 For intCounter = 1 To vsoConnects.Count 
 Set vsoConnect = vsoConnects(intCounter) 
 Set vsoConnectFrom = vsoConnect.FromSheet 
 
 'Print the name of the shape the 
 'Connect object originates from. 
 Debug.Print vsoConnectFrom.Name 
 
 Next intCounter 
 
 Next intCurrentShapeIndex 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
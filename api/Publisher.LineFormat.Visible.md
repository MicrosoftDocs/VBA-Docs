---
title: LineFormat.Visible Property (Publisher)
keywords: vbapb10.chm3408146
f1_keywords:
- vbapb10.chm3408146
ms.prod: publisher
api_name:
- Publisher.LineFormat.Visible
ms.assetid: 508560d2-e143-2d0d-93e7-49141e44b521
ms.date: 06/08/2017
localization_priority: Normal
---


# LineFormat.Visible Property (Publisher)

Returns or sets an  **MsoTriState** constant indicating whether the specified object or the formatting applied to the specified object is visible. Read/write.


## Syntax

 _expression_. **Visible**

 _expression_ A variable that represents a  **LineFormat** object.


## Remarks

The  **Visible** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|Constant|Description|
|:-----|:-----|
| **msoFalse**|The specified object or formatting is not visible.|
| **msoTriStateMixed**|Return value only. The specified shape range contains both objects with visible formatting and objects with invisible formatting.|
| **msoTriStateToggle**| Set value only. Switches the specified object between visible and invisble.|
| **msoTrue**|The specified object or formatting is visible.|

## Example

This example sets the horizontal and vertical offsets for the shadow of shape three on the first page in the active publication. The shadow is offset 5 points to the right of the shape and 3 points above it. If the shape does not already have a shadow, this example adds one to it.


```vb
With ActiveDocument.Pages(1).Shapes(3).Shadow 
 .Visible = msoTrue 
 .OffsetX = 5 
 .OffsetY = -3 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
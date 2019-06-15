---
title: ThreeDFormat.Visible property (Publisher)
keywords: vbapb10.chm3801361
f1_keywords:
- vbapb10.chm3801361
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.Visible
ms.assetid: dbda23fe-de06-cc17-c3bc-2bfb780d3184
ms.date: 06/15/2019
localization_priority: Normal
---


# ThreeDFormat.Visible property (Publisher)

Returns or sets an **[MsoTriState](office.msotristate.md)** constant indicating whether the specified object or the formatting applied to the specified object is visible. Read/write.


## Syntax

_expression_.**Visible**

_expression_ A variable that represents a **[ThreeDFormat](Publisher.ThreeDFormat.md)** object.


## Remarks

The **Visible** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|The specified object or formatting is not visible.|
| **msoTriStateMixed**|Return value only. The specified shape range contains both objects with visible formatting and objects with invisible formatting.|
| **msoTriStateToggle**| Set value only. Switches the specified object between visible and invisible.|
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
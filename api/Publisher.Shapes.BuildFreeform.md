---
title: Shapes.BuildFreeform method (Publisher)
keywords: vbapb10.chm2162723
f1_keywords:
- vbapb10.chm2162723
ms.prod: publisher
api_name:
- Publisher.Shapes.BuildFreeform
ms.assetid: ea24a9a2-e72c-beb3-b17d-161ea41fff1d
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.BuildFreeform method (Publisher)

Builds a freeform object. Returns a **[FreeformBuilder](Publisher.FreeformBuilder.md)** object that represents the freeform as it is being built.


## Syntax

_expression_.**BuildFreeform** (_EditingType_, _X1_, _Y1_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_EditingType_|Required| **[MsoEditingType](office.msoeditingtype.md)**|Specifies the editing type of the first node. Can be one of the **MsoEditingType** constants declared in the Microsoft Office type library.|
|_X1_|Required| **Variant**|The horizontal position of the first node in the freeform drawing relative to the upper-left corner of the page.|
|_Y1_|Required| **Variant**|The vertical position of the first node in the freeform drawing relative to the upper-left corner of the page.|

## Return value

FreeformBuilder


## Example

For the _X1_ and _Y1_ arguments, numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

Use the **[AddNodes](Publisher.FreeformBuilder.AddNodes.md)** method to add segments to the freeform. After you have added at least one segment to the freeform, you can use the **[ConvertToShape](Publisher.FreeformBuilder.ConvertToShape.md)** method to convert the **FreeformBuilder** object into a **Shape** object that has the geometric description that you defined in the **FreeformBuilder** object.

```vb
' Add a new freeform object. 
With ActiveDocument.Shapes _ 
 .BuildFreeform(EditingType:=msoEditingCorner, _ 
 X1:=100, Y1:=100) 
 
 ' Add three more nodes and close the polygon. 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingCorner, _ 
 X1:=200, Y1:=200, X2:=225, Y2:=250, X3:=250, Y3:=200 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingAuto, X1:=200, Y1:=100 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=150, Y1:=50 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=100, Y1:=100 
 
 ' Convert the polygon to a Shape object. 
 .ConvertToShape 
End With 
 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: ShapeRange.LockAspectRatio Property (Publisher)
keywords: vbapb10.chm2293827
f1_keywords:
- vbapb10.chm2293827
ms.prod: publisher
api_name:
- Publisher.ShapeRange.LockAspectRatio
ms.assetid: 8ed4f41f-3395-dd59-29d4-f66afd19ac51
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.LockAspectRatio Property (Publisher)

Returns or sets an  **MsoTriState**constant indicating whether the specified shape retains its original proportions when you resize it. Read/write.


## Syntax

 _expression_. **LockAspectRatio**

 _expression_ A variable that represents a  **ShapeRange** object.


## Remarks

The  **LockAspectRatio** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|Constant|Description|
|:-----|:-----|
| **msoFalse**|The height and width of the shape change independently of one another when you resize it.|
| **msoTriStateMixed**|Return value indicating a combination of  **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTriStateToggle**|Set value that switches between  **msoTrue** and **msoFalse**.|
| **msoTrue**|The specified shape retains its original proportions when you resize it.|

## Example

This example adds a cube to the active publication. The cube can be moved and resized, but not reproportioned.


```vb
Dim shp As Shape 
 
Set shp = ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeCube, _ 
 Left:=50, Top:=50, Width:=100, Height:=200) _ 
 
shp.LockAspectRatio = msoTrue
```



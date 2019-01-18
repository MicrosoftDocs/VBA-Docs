---
title: Master.SpatialSearch Property (Visio)
keywords: vis_sdr.chm10714405
f1_keywords:
- vis_sdr.chm10714405
ms.prod: visio
api_name:
- Visio.Master.SpatialSearch
ms.assetid: d71b05b7-32e1-d3c8-668e-6e96595acd59
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.SpatialSearch Property (Visio)

Returns a  **Selection** object whose shapes meet certain criteria in relation to a point that is expressed in the coordinate space of a page, master, or group. Read-only.


## Syntax

 _expression_. `SpatialSearch`( `_x_` , `_y_` , `_Relation_` , `_Tolerance_` , `_Flags_` )

 _expression_ A variable that represents a [Master](./Visio.Master.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _x_|Required| **Double**|The x-coordinate.|
| _y_|Required| **Double**|The y-coordinate.|
| _Relation_|Required| **Integer**|Any combination of the values of the constants  **visSpatialContainedIn** and **visSpatialTouching**.|
| _Tolerance_|Required| **Double**|A distance in internal drawing units with respect to the coordinate space.|
| _Flags_|Required| **Integer**|Flags that influence the result.|

## Return value

Selection


## Remarks


- The  _relation_ argument can be any combination of the constants defined in **[VisSpatialRelationCodes](Visio.visspatialrelationcodes.md)**. If _relation_ is not specified, the **SpatialSearch** property uses both relationships as criteria.
    
- The  _flags_ argument can be any combination of the values of the constants defined in **[VisSpatialRelationFlags](Visio.visspatialrelationflags.md)** in the Visio type library (except **visSpatialIncludeHidden** , which is reserved for future use, and should not be used).
    
Use the NoShow cell to determine whether a Geometry section is hidden or visible. Hidden Geometry section sections have a value of TRUE and visible Geometry sections have a value of FALSE in the NoShow cell.

Beginning with Microsoft Visio 2002, if  _flags_ contains **visSpatialFrontToBack** , items in the **Selection** object returned by the **SpatialNeighbors** property are ordered front to back. If **visSpatialBackToFront** is set, the items returned are ordered back to front. If this flag is not set, or if you are running an earlier version of Visio, the order is unpredictable. You can determine the order by using the **Index** property of the shapes identified in the **Selection** object _._




 **Note**   When it compares two shapes, the **SpatialSearch** property does not consider the width of a shape's line, shadows, line ends, control points, or connection points.



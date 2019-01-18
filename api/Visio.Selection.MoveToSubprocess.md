---
title: Selection.MoveToSubprocess Method (Visio)
keywords: vis_sdr.chm11162210
f1_keywords:
- vis_sdr.chm11162210
ms.prod: visio
api_name:
- Visio.Selection.MoveToSubprocess
ms.assetid: a61f1e93-06a3-6ddc-8cae-f92212078c96
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.MoveToSubprocess Method (Visio)

Moves the selection to the specified page, and drops a replacement shape on the source page and links it to the target page. Returns the selection of moved shapes on the target page.


## Syntax

 _expression_. `MoveToSubprocess`( `_Page_` , `_ObjectToDrop_` , `_NewShape_` )

 _expression_ A variable that represents a '[Selection](Visio.Selection.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Page_|Required| **[Page](Visio.Page.md)**|The subprocess page to which the selection should be moved. You cannot pass the current page.|
| _ObjectToDrop_|Required| **[UNKNOWN]**|The replacement shape to drop.|
| _NewShape_|Optional| **[Shape](Visio.Shape.md)**|Out parameter. Returns the shape that is linked to the new page.|

## Return value

 **Selection**


## Remarks

The  _ObjectToDrop_ parameter is typically a Microsoft Visio object, such as a **[Master](Visio.Master.md)** or **[MasterShortcut](Visio.MasterShortcut.md)** object. However, it can be any OLE object that provides an **IDataObject** interface. If _ObjectToDrop_ is null, Visio drops a default shape.



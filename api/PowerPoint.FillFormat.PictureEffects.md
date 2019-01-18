---
title: FillFormat.PictureEffects Property (PowerPoint)
keywords: vbapp10.chm552033
f1_keywords:
- vbapp10.chm552033
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.PictureEffects
ms.assetid: 01897ad5-84c9-f98e-8c2f-9a9e5c13bc2e
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.PictureEffects Property (PowerPoint)

Returns an object that represents the picture or texture fill for the specified fill format. Read-only.


## Syntax

 _expression_. `PictureEffects`

 _expression_ A variable that represents a [FillFormat](./PowerPoint.FillFormat.md) object.


## Return value

[PictureEffects](Office.PictureEffects.md)


## Remarks

A picture or texture fill can be specified in the formatting for various elements (shapes) in a chart. For example, you can use the  **Format Data Series** dialog box to format the columns in a **Column** chart to a picture or texture fill. In this case, the **PictureEffects** property returns a **PictureEffects** collection that corresponds to the settings associated with the **Picture or texture fill** option in the **Fill** category of the **Format Data Series** dialog box.


## See also


[FillFormat Object](PowerPoint.FillFormat.md)


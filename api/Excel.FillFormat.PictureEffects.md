---
title: FillFormat.PictureEffects property (Excel)
ms.prod: excel
api_name:
- Excel.FillFormat.PictureEffects
ms.assetid: bb5e8d9d-a878-c8c4-b198-ef7269f837f0
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.PictureEffects property (Excel)

Returns an object that represents the picture or texture fill for the specified fill format. Read-only


## Syntax

_expression_. `PictureEffects`

_expression_ A variable that repressents a '[FillFormat](Excel.FillFormat.md)' object.


## Return value

[PictureEffects](Office.PictureEffects.md)


## Remarks

A picture or texture fill can be specified in the formatting for various elements (shapes) in a chart. For example, you can use the  **Format Data Series** dialog box to format the columns in a **Column** chart to a picture or texture fill. In this case, the **PictureEffects** property returns a **PictureEffects** collection that corresponds to the settings associated with the **Picture or texture fili** option in the **Fill** category of the **Format Data Series** dialog box.


## See also


[FillFormat Object](Excel.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
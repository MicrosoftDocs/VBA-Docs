---
title: FillFormat.RotateWithObject property (Publisher)
keywords: vbapb10.chm2359585
f1_keywords:
- vbapb10.chm2359585
ms.prod: publisher
ms.assetid: a1e5f826-4200-4eac-204d-f17717160f33
ms.date: 06/07/2019
localization_priority: Normal
---


# FillFormat.RotateWithObject property (Publisher)

Returns or sets whether the fill rotates with the specified shape. Read/write.


## Syntax

_expression_.**RotateWithObject**

_expression_ A variable that represents a **[FillFormat](publisher.fillformat.md)** object.


## Return value

**[MsoTriState](Office.MsoTriState.md)**


## Remarks

The value returned by the **RotateWithObject** property can be one of the **MsoTriState** constants listed in the following table.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|The fill does not rotate with the shape.|
| **msoTrue**|The fill rotates with the shape.|

The setting of the **RotateWithObject** property corresponds to the setting of the **Rotate with shape** box in the **Fill** pane of the **Format Shape** dialog box in the Publisher user interface.

> [!NOTE] 
> The **Rotate with shape** box appears only if you have either the **Gradient fill** or **Picture or texture fill** option buttons selected in the **Fill** pane of the **Format Shape** dialog box.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
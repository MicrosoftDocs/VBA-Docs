---
title: FillFormat.RotateWithObject Property (Publisher)
keywords: vbapb10.chm2359585
f1_keywords:
- vbapb10.chm2359585
ms.prod: publisher
ms.assetid: a1e5f826-4200-4eac-204d-f17717160f33
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.RotateWithObject Property (Publisher)

Returns or sets whether the fill rotates with the specified shape. Read/write.


## Syntax

 _expression_. **RotateWithObject**

 _expression_ A variable that represents a  **FillFormat** object.


## Return value

 **MSOTRISTATE**


## Remarks

The value returned by the  **RotateWithObject** property can be one of the ** [MsoTriState](Office.MsoTriState.md)** constants listed in the following table.



|Constant|Description|
|:-----|:-----|
| **msoFalse**|The fill does not rotate with the shape.|
| **msoTrue**|The fill rotates with the shape.|

The setting of the  **RotateWithObject** property corresponds to the setting of the **Rotate with shape** box on the **Fill** pane of the **Format Shape** dialog box in the Publisher user interface.


> [!NOTE] 
> The  **Rotate with shape** box appears only if you have either the **Gradient fill** or **Picture or texture fill** option buttons selected on the **Fill** pane of the **Format Shape** dialog box.


## See also


 [FillFormat Object](Publisher.FillFormat.md)


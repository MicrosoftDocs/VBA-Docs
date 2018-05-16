---
title: Slicer.DisableMoveResizeUI Property (Excel)
keywords: vbaxl10.chm905077
f1_keywords:
- vbaxl10.chm905077
ms.prod: excel
api_name:
- Excel.Slicer.DisableMoveResizeUI
ms.assetid: 2477e495-e61a-6981-6df2-5bb1cb480576
ms.date: 06/08/2017
---


# Slicer.DisableMoveResizeUI Property (Excel)

Returns or sets whether the specified slicer can be moved or resized by using the user interface. Read/write.


## Syntax

 _expression_ . **DisableMoveResizeUI**

 _expression_ A variable that represents a **[Slicer](Excel.Slicer.md)** object.


### Return Value

Boolean


## Remarks

 **True** if the slicer cannot be moved or resized by selecting borders or handles in the user interface; otherwise **False** . The default value is **False** . Setting the **DisableMoveResizeUI** property to **True** affects only the user interface. Moving or resizing the slicer by setting properties such as the **[Top](Excel.Slicer.Top.md)** , **[Left](Excel.Slicer.Left.md)** , **[Width](Excel.Slicer.Width.md)** , or **[Height](Excel.Slicer.Height.md)** properties of the **Slicer** object from code is not disabled.


## See also


#### Concepts


[Slicer Object](Excel.Slicer.md)


---
title: Slicer.DisableMoveResizeUI property (Excel)
keywords: vbaxl10.chm905077
f1_keywords:
- vbaxl10.chm905077
ms.prod: excel
api_name:
- Excel.Slicer.DisableMoveResizeUI
ms.assetid: 2477e495-e61a-6981-6df2-5bb1cb480576
ms.date: 05/16/2019
localization_priority: Normal
---


# Slicer.DisableMoveResizeUI property (Excel)

Returns or sets whether the specified slicer can be moved or resized by using the user interface. Read/write.


## Syntax

_expression_.**DisableMoveResizeUI**

_expression_ A variable that represents a **[Slicer](Excel.Slicer.md)** object.


## Return value

**Boolean**


## Remarks

**True** if the slicer cannot be moved or resized by selecting borders or handles in the user interface; otherwise, **False**. The default value is **False**. 

Setting the **DisableMoveResizeUI** property to **True** affects only the user interface. 

Moving or resizing the slicer by setting properties, such as **[Top](Excel.Slicer.Top.md)**, **[Left](Excel.Slicer.Left.md)**, **[Width](Excel.Slicer.Width.md)**, or **[Height](Excel.Slicer.Height.md)**, from code is not disabled.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
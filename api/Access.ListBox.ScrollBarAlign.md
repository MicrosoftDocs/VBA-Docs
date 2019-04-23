---
title: ListBox.ScrollBarAlign property (Access)
keywords: vbaac10.chm11293
f1_keywords:
- vbaac10.chm11293
ms.prod: access
api_name:
- Access.ListBox.ScrollBarAlign
ms.assetid: 6eb9b2d1-e306-5980-7ad0-ff0b9c1cd0c6
ms.date: 03/02/2019
localization_priority: Normal
---


# ListBox.ScrollBarAlign property (Access)

You can use the **ScrollBarAlign** property to specify or determine the alignment of a vertical scroll bar. Read/write **Byte**.


## Syntax

_expression_.**ScrollBarAlign**

_expression_ A variable that represents a **[ListBox](Access.ListBox.md)** object.


## Remarks

The **ScrollBarAlign** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|System|0|A vertical scroll bar is placed on the left if the form or report **Orientation** property is right to left, and on the right if the form or report **Orientation** property is left to right.|
|Right|1|Aligns the vertical scroll bar on the right side of the control.|
|Left|2|Aligns the vertical scroll bar on the left side of the control.|

For combo and list boxes, **ScrollBarAlign** also controls the placement of the box button above the scroll bar.


## Example

The following example aligns the vertical scroll bar on the left side of the **Country** combo box in the **International Shipping** form.


```vb
Forms("International Shipping").Controls("Country").ScrollBarAlign = 2
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
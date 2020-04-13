---
title: ScrollBar.SmallChange Property (Outlook Forms Script)
keywords: olfm10.chm2001940
f1_keywords:
- olfm10.chm2001940
ms.prod: outlook
ms.assetid: cd8b6b7f-118a-1cda-00af-11ab74f6617a
ms.date: 06/08/2017
localization_priority: Normal
---


# ScrollBar.SmallChange Property (Outlook Forms Script)

Returns or sets an **Integer** that specifies the amount of movement that occurs when the user clicks either scroll arrow in a **[ScrollBar](Outlook.scrollbar.md)**. Read/write.


## Syntax

_expression_.**SmallChange**

_expression_ A variable that represents a **ScrollBar** object.


## Remarks

The **SmallChange** property specifies the amount of change to the **[Value](Outlook.scrollbar.value.md)** property.

The **SmallChange** property does not have units.

Any integer is an acceptable setting for this property. The recommended range of values is from -32,767 to +32,767. The default value is 1.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
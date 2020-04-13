---
title: CheckBox.MousePointer Property (Outlook Forms Script)
keywords: olfm10.chm2001550
f1_keywords:
- olfm10.chm2001550
ms.prod: outlook
ms.assetid: 7787fce4-564a-ad9e-6e54-d4cd6a0a3e8a
ms.date: 06/08/2017
localization_priority: Normal
---


# CheckBox.MousePointer Property (Outlook Forms Script)

Returns or sets an **Integer** that specifies the type of pointer displayed when the user positions the mouse over a particular object. Read/write.


## Syntax

_expression_.**MousePointer**

_expression_ A variable that represents a **CheckBox** object.


## Remarks

The settings for  **MousePointer** are:



|Value|Description|
|:-----|:-----|
|0|Standard pointer. The image is determined by the object (default).|
|1|Arrow.|
|2|Cross-hair pointer.|
|3|I-beam.|
|6|Double arrow pointing northeast and southwest.|
|7|Double arrow pointing north and south.|
|8|Double arrow pointing northwest and southeast.|
|9|Double arrow pointing west and east.|
|10|Up arrow.|
|11|Hourglass.|
|12|"Not" symbol (circle with a diagonal line) on top of the object being dragged. Indicates an invalid drop target.|
|13|Arrow with an hourglass.|
|14|Arrow with a question mark.|
|15|Size all cursor (arrows pointing north, south, east, and west).|
|99|Uses the icon specified by the  **[MouseIcon](Outlook.checkbox.mouseicon.md)** property.|

Use the  **MousePointer** property when you want to indicate changes in functionality as the mouse pointer passes over controls on a form. For example, the hourglass setting (11) is useful to indicate that the user must wait for a process or operation to finish.

Some icons vary depending on system settings, such as the icons associated with desktop themes.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
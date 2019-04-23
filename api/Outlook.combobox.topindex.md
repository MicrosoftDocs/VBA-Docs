---
title: ComboBox.TopIndex Property (Outlook Forms Script)
keywords: olfm10.chm2002120
f1_keywords:
- olfm10.chm2002120
ms.prod: outlook
ms.assetid: e5fcb92e-5f0c-2dc5-c074-022174a0b4e7
ms.date: 06/08/2017
localization_priority: Normal
---


# ComboBox.TopIndex Property (Outlook Forms Script)

Returns or sets a  **Long** that represents the index of the item displayed in the topmost position in the list portion of the **[ComboBox](Outlook.combobox.md)**. Read/write.


## Syntax

_expression_.**TopIndex**

_expression_ A variable that represents a  **ComboBox** object.


## Remarks

The default is 0, which identifies the first item in the list.

Returns the value -1 if the list is empty or not displayed.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
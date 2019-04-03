---
title: ComboBox.ListWidth Property (Outlook Forms Script)
keywords: olfm10.chm2001460
f1_keywords:
- olfm10.chm2001460
ms.prod: outlook
ms.assetid: bc16a1c0-5db3-64a3-21d3-c1537052aa2b
ms.date: 06/08/2017
localization_priority: Normal
---


# ComboBox.ListWidth Property (Outlook Forms Script)

Returns or sets a  **Variant** that specifies the width of the list in a **[ComboBox](Outlook.combobox.md)**. Read/write.


## Syntax

_expression_.**ListWidth**

_expression_ A variable that represents a  **ComboBox** object.


## Remarks

A value of zero makes the list as wide as the  **ComboBox**. The default value is to make the list as wide as the text portion of the control.

If you want to display a multicolumn list, enter a value that will make the list box wide enough to fit all the columns.

 When designing combo boxes, be sure to leave enough space to display your data and for a vertical scroll bar.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
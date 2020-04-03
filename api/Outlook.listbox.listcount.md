---
title: ListBox.ListCount Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 1a06637a-8c23-e7a5-f7e4-7a04dcb227fc
ms.date: 06/08/2017
localization_priority: Normal
---


# ListBox.ListCount Property (Outlook Forms Script)

Returns a **Long** that represents the number of list entries in a control. Read-only.


## Syntax

_expression_.**ListCount**

_expression_ A variable that represents a **ListBox** object.


## Remarks

 **ListCount** is the number of rows over which you can scroll. **ListCount** is always one greater than the largest value for the **[ListIndex](Outlook.listbox.listindex.md)** property, because index numbers begin with 0 and the count of items begins with 1. If no item is selected, **ListCount** is 0 and **ListIndex** is -1.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: ComboBox.AutoWordSelect Property (Outlook Forms Script)
keywords: olfm10.chm2000760
f1_keywords:
- olfm10.chm2000760
ms.prod: outlook
ms.assetid: 721086f4-2400-31c1-9b32-0e7100a5c78a
ms.date: 06/08/2017
localization_priority: Normal
---


# ComboBox.AutoWordSelect Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether the basic unit used to extend a selection is a word or a single character. Read/write.


## Syntax

_expression_.**AutoWordSelect**

_expression_ A variable that represents a  **ComboBox** object.


## Remarks

 **True** if uses a word as the basic unit (default), **False** if uses a character as the basic unit.

The  **AutoWordSelect** property specifies how the selection extends or contracts in the edit region of a **[ComboBox](Outlook.combobox.md)**.

If the user places the insertion point in the middle of a word and then extends the selection while  **AutoWordSelect** is **True**, the selection includes the entire word.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
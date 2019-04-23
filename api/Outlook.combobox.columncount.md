---
title: ComboBox.ColumnCount Property (Outlook Forms Script)
keywords: olfm10.chm2000940
f1_keywords:
- olfm10.chm2000940
ms.prod: outlook
ms.assetid: 9bbdcdfa-25c8-5113-8532-6bf4857aef67
ms.date: 06/08/2017
localization_priority: Normal
---


# ComboBox.ColumnCount Property (Outlook Forms Script)

Returns or sets a  **Long** that represents the number of columns to display in a combo box. Read/write.


## Syntax

_expression_.**ColumnCount**

_expression_ A variable that represents a  **ComboBox** object.


## Remarks

Setting  **ColumnCount** to 0 displays zero columns, and setting it to -1 displays all the available columns. For an unbound data source, there is a 10-column limit (0 to 9).

You can use the  **[ColumnWidths](Outlook.combobox.columnwidths.md)** property to set the width of the columns displayed in the control.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
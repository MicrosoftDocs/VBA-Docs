---
title: ComboBox.List Property (Outlook Forms Script)
keywords: olfm10.chm2001400
f1_keywords:
- olfm10.chm2001400
ms.prod: outlook
ms.assetid: 687f44e8-7b4b-eab5-93b8-022cd4d1c302
ms.date: 06/08/2017
localization_priority: Normal
---


# ComboBox.List Property (Outlook Forms Script)

Returns or sets a  **Variant** that represents the specified entry in a **[ComboBox](Outlook.combobox.md)**. Read/write.


## Syntax

_expression_.**List**(**_pvargIndex_**,  **_pvargColumn_**)

_expression_ A variable that represents a  **ComboBox** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|pvargIndex|Optional| **Variant**|An integer with a range from 0 to one less than the number of entries in the list of the  **ComboBox**.|
|pvargColumn|Optional| **Variant**|An integer with a range from 0 to one less than the number of columns in the list of the  **ComboBox**.|

## Remarks

Row and column numbering begins with zero. That is, the row number of the first row in the list is zero; the column number of the first column is zero. The number of the second row or column is 1, and so on.

The  **List** property works with the **[ListCount](Outlook.combobox.listcount.md)** and **[ListIndex](Outlook.combobox.listindex.md)** properties. Use **List** to access list items. A list is a variant array; each item in the list has a row number and a column number.

Initially, a  **ComboBox** contains an empty list.

To specify items you want to display in a  **ComboBox**, use the  **[AddItem](Outlook.combobox.additem.md)** method. To remove items, use the **[RemoveItem](Outlook.combobox.removeitem.md)** method.

Use  **List** to copy an entire two-dimensional array of values to a control. Use **AddItem** to load a one-dimensional array or to load an individual element.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

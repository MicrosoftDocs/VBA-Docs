---
title: ControlFormat.MultiSelect property (Excel)
keywords: vbaxl10.chm630087
f1_keywords:
- vbaxl10.chm630087
ms.prod: excel
api_name:
- Excel.ControlFormat.MultiSelect
ms.assetid: 5ec1e5b6-37ab-465b-bf81-4955f6fd0f31
ms.date: 04/23/2019
localization_priority: Normal
---


# ControlFormat.MultiSelect property (Excel)

Returns or sets the selection mode of the specified list box. Can be one of the following [constants](excel.constants.md): **xlNone**, **xlSimple**, or **xlExtended**. Read/write **Long**.


## Syntax

_expression_.**MultiSelect**

_expression_ A variable that represents a **[ControlFormat](Excel.ControlFormat.md)** object.


## Remarks

Single select (**xlNone**) allows only one item at a time to be selected. Choosing the mouse or pressing the Spacebar cancels the selection and selects the chosen item.

Simple multiselect (**xlSimple**) toggles the selection on an item in the list when it's chosen with the mouse or the Spacebar is pressed when the focus is on the item. This mode is appropriate for pick lists, in which there are often multiple items selected.

Extended multiselect (**xlExtended**) usually acts like a single-selection list box, so when you choose an item, you cancel all other selections. When you hold down Shift while choosing the mouse or pressing an arrow key, you select items sequentially from the current item. When you hold down Ctrl while choosing the mouse, you add single items to the list. This mode is appropriate when multiple items are allowed but not often used.

You can use the **Value** or **ListIndex** property to return and set the selected item in a single-select list box.

You cannot link multiselect list boxes by using the **LinkedCell** property.


## Example

This example creates a simple multiselect list box.

```vb
Set lb = Worksheets(1).Shapes.AddFormControl(xlListBox, _ 
 Left:=10, Top:=10, Height:=100, Width:100) 
lb.ControlFormat.MultiSelect = xlSimple
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
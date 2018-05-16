---
title: EmptyCell.StatusBarText Property (Access)
keywords: vbaac10.chm14341
f1_keywords:
- vbaac10.chm14341
ms.prod: access
api_name:
- Access.EmptyCell.StatusBarText
ms.assetid: c0c02302-b2a5-ab77-8f4e-26c0402bb4b0
ms.date: 06/08/2017
---


# EmptyCell.StatusBarText Property (Access)

You can use the  **StatusBarText** property to specify the text that is displayed in the status bar when a control is selected. Read/write **String**.


## Syntax

 _expression_. **StatusBarText**

 _expression_ A variable that represents an **EmptyCell** object.


## Remarks

You set the  **StatusBarText** property by using a string expression up to 255 characters long.


 **Note**  The length of the text you can display in the status bar depends on your computer hardware and video display.

You can use the  **StatusBarText** property to provide specific information about a control. For example, when a text box has the focus, a brief instruction can tell the user what kind of data to enter


 **Note**  You can also use the  **ControlTipText** property to display a ScreenTip for a control.

If you create a control by dragging a field from the field list, the value in a field's  **Description** property is copied to the **StatusBarText** property.


## Example

The following example sets the status bar text to be displayed when the "Address_TextBox" control in the "Mailing List" form has the focus in Form View. 


```vb
Forms("Mailing List").Controls("Address_TextBox"). _ 
 StatusBarText = "Enter the company's mailing address." 

```


## See also


#### Concepts


[EmptyCell Object](Access.EmptyCell.md)


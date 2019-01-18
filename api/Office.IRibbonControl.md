---
title: IRibbonControl object (Office)
keywords: vbaof11.chm288000
f1_keywords:
- vbaof11.chm288000
ms.prod: office
api_name:
- Office.IRibbonControl
ms.assetid: 63aef709-e1d3-b1a6-76af-b568ad0e69ae
ms.date: 01/16/2019
localization_priority: Normal
---


# IRibbonControl object (Office)

Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.


## Remarks

The **IRibbonControl** object contains the name (ID) of the control and the current **Window** object for the Ribbon UI control.


## Example

The following example, written in C#, shows two procedures called from the **onAction** event procedure of a **Button** control and a **ToggleButton** control. 

In the first procedure, the **IRibbonControl** object representing the control is passed into the procedure, and a message box is displayed indicating that the button was pressed along with the ID of the button. 

The second procedure is similar to the first with the addition of a **Boolean** parameter indicating that the button was pressed.


```cs
public void ButtonOnAction(IRibbonControl control) 
{ 
 MessageBox.Show("Button clicked: " + control.Id); 
} 
 
public void ToggleButtonOnAction(IRibbonControl control, bool pressed) 
{ 
...if (pressed) 
 MessageBox.Show("ToggleButton was switched on."); 
 else 
 MessageBox.Show("ToggleButton was switched off."); 
}
```


## See also

- [IRibbonControl object members](overview/library-reference/iribboncontrol-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
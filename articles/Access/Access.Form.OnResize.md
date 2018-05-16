---
title: Form.OnResize Property (Access)
keywords: vbaac10.chm13442
f1_keywords:
- vbaac10.chm13442
ms.prod: access
api_name:
- Access.Form.OnResize
ms.assetid: 84e6df44-53d2-19c9-e8c5-47681649c7e8
ms.date: 06/08/2017
---


# Form.OnResize Property (Access)

Sets or returns the value of the  **On Resize** box in the **Properties** window of a form. Read/write **String**.


## Syntax

 _expression_. **OnResize**

 _expression_ A variable that represents a **Form** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **Resize** event occurs when a form is opened and whenever the size of a form changes.

The  **OnResize** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Resize** box in the form's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Resize** box is blank, the property value is an empty string.


## Example

The following example prints the value of the  **OnResize** property in the Immediate window for the "Order Entry" form.


```vb
Debug.Print Forms("Order Entry").OnResize
```


## See also


#### Concepts


[Form Object](Access.Form.md)


---
title: Form.OnUnload Property (Access)
keywords: vbaac10.chm13443
f1_keywords:
- vbaac10.chm13443
ms.prod: access
api_name:
- Access.Form.OnUnload
ms.assetid: 70544311-921c-a610-6fbe-bd3bbef0a6a5
ms.date: 06/08/2017
---


# Form.OnUnload Property (Access)

Sets or returns the value of the  **On Unload** box in the **Properties** window of a form. Read/write **String**.


## Syntax

 _expression_. **OnUnload**

 _expression_ A variable that represents a **Form** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **Unload** event occurs after a form is closed but before it's removed from the screen.

The  **OnUnload** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Unload** box in the form's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Unload** box is blank, the property value is an empty string.


## Example

The following example prints the value of the  **OnUnload** property in the Immediate window for the "Order Entry" form.


```vb
Debug.Print Forms("Order Entry").OnUnload
```


## See also


#### Concepts


[Form Object](Access.Form.md)


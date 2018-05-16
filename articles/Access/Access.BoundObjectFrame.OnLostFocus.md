---
title: BoundObjectFrame.OnLostFocus Property (Access)
keywords: vbaac10.chm10967
f1_keywords:
- vbaac10.chm10967
ms.prod: access
api_name:
- Access.BoundObjectFrame.OnLostFocus
ms.assetid: 78ee2d7f-89d4-e9d2-a0ce-ecd6d35a98c3
ms.date: 06/08/2017
---


# BoundObjectFrame.OnLostFocus Property (Access)

Sets or returns the value of the  **On Lost Focus** box in the **Properties** window of the specified object. Read/write **String**.


## Syntax

 _expression_. **OnLostFocus**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **LostFocus** event occurs when the object loses the focus.

The  **OnLostFocus** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Lost Focus** box in the object's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Lost Focus** box is blank, the property value is an empty string.


## Example

The following example prints the value of the  **OnLostFocus** property in the Immediate window for the button named "OK" on the "Order Entry" form.


```vb
Debug.Print Forms("Order Entry").Controls("OK").OnLostFocus
```


## See also


#### Concepts


[BoundObjectFrame Object](Access.BoundObjectFrame.md)


---
title: ToggleButton.OnMouseDown property (Access)
keywords: vbaac10.chm11755
f1_keywords:
- vbaac10.chm11755
ms.prod: access
api_name:
- Access.ToggleButton.OnMouseDown
ms.assetid: a9bbf8a5-4e62-fa9e-63a4-2f59cd2734f4
ms.date: 02/23/2019
localization_priority: Normal
---


# ToggleButton.OnMouseDown property (Access)

Sets or returns the value of the **On Mouse Down** box in the Properties window. Read/write **String**.


## Syntax

_expression_.**OnMouseDown**

_expression_ A variable that represents a **[ToggleButton](Access.ToggleButton.md)** object.


## Remarks

This property is helpful for programmatically changing the action that Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The **MouseDown** event occurs when the user clicks the mouse button while the mouse pointer rests over the object.

The **OnMouseDown** value will be one of the following, depending on the selection chosen in the Choose Builder window (accessed by choosing the **Build** button next to the **On Mouse Down** box in the object's Properties window):

- If you choose Expression Builder, the value will be =_expression_, where _expression_ is the expression from the Expression Builder window.
    
- If you choose Macro Builder, the value is the name of the macro. 
    
- If you choose Code Builder, the value will be [Event Procedure]. 
    
If the **On Mouse Down** box is blank, the property value is an empty string.


## Example

The following example prints the value of the **OnMouseDown** property in the Immediate window for the button named **OK** on the **Order Entry** form.

```vb
Debug.Print Forms("Order Entry").Controls("OK").OnMouseDown
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
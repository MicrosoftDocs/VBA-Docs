---
title: Report.OnDblClick Property (Access)
keywords: vbaac10.chm13862
f1_keywords:
- vbaac10.chm13862
ms.prod: access
api_name:
- Access.Report.OnDblClick
ms.assetid: b92a1b2b-4f27-4f45-959c-6f1aec557004
ms.date: 06/08/2017
---


# Report.OnDblClick Property (Access)

Sets or returns the value of the  **On Dbl Click** box in the **Properties** window. Read/write **String**.


## Syntax

 _expression_. **OnDblClick**

 _expression_ A variable that represents a **Report** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **DblClick** event occurs when a user presses and releases the left mouse button twice over an object within the double-click time limit of the system.

The  **OnDblClick** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Dbl Click** box in the report's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Dbl Click** box is blank, the property value is an empty string.


## Example

The following example associates the  **DblClick** event with the "OK_DblClick" event procedure for the button named "OK" on the "Catalog" report, if there is currently no association.


```vb
With Reports("Catalog").Controls("OK") 
 If .OnDblClick = "" Then 
 .OnDblClick = "[Event Procedure]" 
 End If 
End With 

```


## See also


#### Concepts


[Report Object](Access.Report.md)


---
title: Report.OnApplyFilter Property (Access)
keywords: vbaac10.chm13870
f1_keywords:
- vbaac10.chm13870
ms.prod: access
api_name:
- Access.Report.OnApplyFilter
ms.assetid: 18e5b016-19a0-46bb-c552-c4bb8d458ca4
ms.date: 06/08/2017
---


# Report.OnApplyFilter Property (Access)

Sets or returns the value of the  **On Apply Filter** box in the **Properties** window of a report. Read/write **String**.


## Syntax

 _expression_. **OnApplyFilter**

 _expression_ A variable that represents a **Report** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **ApplyFilter** event occurs when a filter is applied or removed.

The  **OnApplyFilter** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Apply Filter** box in the report's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Apply Filter** box is blank, the property value is an empty string.


## Example

The following example associates the  **OnApplyFilter** property for the "Catalog" report to the event "Report_ApplyFilter".


```vb
Reports("Catalog").OnApplyFilter = "[Event Procedure]"
```


## See also


#### Concepts


[Report Object](Access.Report.md)


---
title: Report.ApplyFilter event (Access)
keywords: vbaac10.chm13899
f1_keywords:
- vbaac10.chm13899
ms.prod: access
api_name:
- Access.Report.ApplyFilter
ms.assetid: 46cbe83d-4395-d9e6-3187-c51152269e62
ms.date: 03/08/2019
localization_priority: Normal
---


# Report.ApplyFilter event (Access)

Occurs when a filter is applied to a report.


## Syntax

_expression_.**ApplyFilter** (_Cancel_, _ApplyType_)

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|The setting determines if the **ApplyFilter** event occurs. Setting the _Cancel_ argument to **True** cancels the **ApplyFilter** event and the filter is not applied to the report.|
| _ApplyType_|Required|**Integer**|Returns the type of filter that was applied.|

## Remarks

To run a macro or event procedure when this event occurs, set the **[OnApplyFilter](Access.Report.OnApplyFilter.md)** property to the name of the macro or to [Event Procedure].

You can use the **ApplyFilter** event to:

- Make sure that the filter that is being applied is correct. For example, you may want to be sure that any filter applied to an **Orders** report includes criteria restricting the **OrderDate** field. To do this, check the report's **[Filter](Access.Report.Filter(property).md)** or **[ServerFilter](Access.Report.ServerFilter.md)** property value to make sure that this criteria is included in the WHERE clause expression.
    
- Change the display of the report before the filter is applied. For example, when you apply a certain filter, you may want to disable or hide some fields that aren't appropriate for the records displayed by this filter.
    
- Undo or change actions that you took when the **Filter** event occurred. For example, you can disable or hide some controls on the report when the user is creating the filter because you don't want these controls to be included in the filter criteria. You can then enable or show these controls after the filter is applied.
    
The actions in the **ApplyFilter** event procedure or macro occur before the filter is applied or removed, or after the Advanced Filter/Sort window is closed, but before the report is redisplayed. The criteria you've entered in the newly created filter are available to the **ApplyFilter** event procedure or macro as the setting of the **Filter** or **ServerFilter** property.

The **ApplyFilter** event does not occur when the user does one of the following:

- Applies or removes a filter by using the ApplyFilter, OpenReport, or ShowAllRecords actions in a macro, or their corresponding methods of the **DoCmd** object in Visual Basic.
    
- Uses the Close action or the **Close** method of the **DoCmd** object to close the Advanced Filter/Sort window.
    
- Sets the **Filter** or **ServerFilter** property or **FilterOn** property in a macro or Visual Basic (although you can set these properties in an **ApplyFilter** event procedure or macro).
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
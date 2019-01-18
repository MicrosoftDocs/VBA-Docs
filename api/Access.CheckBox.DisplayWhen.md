---
title: CheckBox.DisplayWhen property (Access)
keywords: vbaac10.chm10702
f1_keywords:
- vbaac10.chm10702
ms.prod: access
api_name:
- Access.CheckBox.DisplayWhen
ms.assetid: 9236d99e-df4d-5342-e60c-162abe7de8d6
ms.date: 06/08/2017
localization_priority: Normal
---


# CheckBox.DisplayWhen property (Access)

You can use the  **DisplayWhen** property to specify which of a form's controls you want displayed on screen and in print. Read/write **Byte**.


## Syntax

_expression_. `DisplayWhen`

_expression_ A variable that represents a [CheckBox](Access.CheckBox.md) object.


## Remarks

The  **DisplayWhen** property applies only to the following form sections: detail, form header, and form footer. It also applies to all controls (except page breaks) on a form.

The  **DisplayWhen** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Always|0|(Default) The object appears in Form view and when printed.|
|Print Only|1|The object is hidden in Form view but appears when printed.|
|Screen Only|2|The object appears in Form view but not when printed.|

For controls, you can set the default for this property by using the default control style or the  **DefaultControl** property in Visual Basic.

In many cases, certain controls are useful only in Form view. To prevent Microsoft Access from printing these controls, you can set their  **DisplayWhen** property to Screen Only. For example, you might have a command button or instructions on a form that you don't want printed. Or you might have form header and form footer sections that you don't want displayed on screen but that you do want printed. In this case, you should set the **DisplayWhen** property to Print Only.

For reports, use the  **Format** and **Retreat** events to specify an event procedure or macro that sets the **Visible** property of controls you don't want printed. You can also cancel the Format or **Print** event for a report section to prevent the section from being printed.


## See also


[CheckBox Object](Access.CheckBox.md)


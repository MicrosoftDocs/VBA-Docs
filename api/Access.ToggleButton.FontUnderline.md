---
title: ToggleButton.FontUnderline property (Access)
keywords: vbaac10.chm11727
f1_keywords:
- vbaac10.chm11727
ms.prod: access
api_name:
- Access.ToggleButton.FontUnderline
ms.assetid: fef06d9f-f21f-a753-9822-f1e823ab10b4
ms.date: 06/08/2017
localization_priority: Normal
---


# ToggleButton.FontUnderline property (Access)

You can use the  **FontUnderline** property to specify whether text is underlined in the following situations:


- When displaying or printing controls on forms and reports.
    
- When using the  **Print** method on a report.
    

 Read/write **Boolean**.


## Syntax

_expression_. `FontUnderline`

_expression_ A variable that represents a [ToggleButton](Access.ToggleButton.md) object.


## Remarks

The  **FontUnderline** property uses the following settings.



|Setting|Description|
|:-----|:-----|
|**True**|The text is underlined.|
|**False**|(Default) The text isn't underlined.|

For reports, you can use this property only in an event procedure or in a macro specified by the  **OnPrint** event property setting.

You can set the default for this property by using the default control style or the  **DefaultControl** property in Visual Basic.


## See also


[ToggleButton Object](Access.ToggleButton.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
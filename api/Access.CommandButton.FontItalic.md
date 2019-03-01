---
title: CommandButton.FontItalic property (Access)
keywords: vbaac10.chm10475
f1_keywords:
- vbaac10.chm10475
ms.prod: access
api_name:
- Access.CommandButton.FontItalic
ms.assetid: a82d5e83-b892-a006-e68a-cda3c2c82d1d
ms.date: 03/01/2019
localization_priority: Normal
---


# CommandButton.FontItalic property (Access)

You can use the **FontItalic** property to specify whether text is italic in the following situations:

- When displaying or printing controls on forms and reports.    
- When using the **[Print](Access.Report.Print.md)** method on a report.
    
Read/write **Boolean**.


## Syntax

_expression_.**FontItalic**

_expression_ A variable that represents a **[CommandButton](Access.CommandButton.md)** object.


## Remarks

The **FontItalic** property uses the following settings.

|Setting|Description|
|:-----|:-----|
|**True**|The text is italic.|
|**False**|(Default) The text isn't italic.|

For reports, you can use this property only in an event procedure or in a macro specified by the **OnPrint** event property setting.

You can set the default for this property by using the default control style or the **[DefaultControl](access.form.defaultcontrol.md)** property in Visual Basic.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
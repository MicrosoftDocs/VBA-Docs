---
title: TabControl.FontItalic property (Access)
keywords: vbaac10.chm12093
f1_keywords:
- vbaac10.chm12093
ms.prod: access
api_name:
- Access.TabControl.FontItalic
ms.assetid: 7147f297-b88b-c0b0-94eb-73977c1809ba
ms.date: 03/01/2019
localization_priority: Normal
---


# TabControl.FontItalic property (Access)

You can use the **FontItalic** property to specify whether text is italic in the following situations:

- When displaying or printing controls on forms and reports.    
- When using the **[Print](Access.Report.Print.md)** method on a report.
    
Read/write **Boolean**.


## Syntax

_expression_.**FontItalic**

_expression_ A variable that represents a **[TabControl](Access.TabControl.md)** object.


## Remarks

The **FontItalic** property uses the following settings.

|Setting|Description|
|:-----|:-----|
|**True**|The text is italic.|
|**False**|(Default) The text isn't italic.|

For reports, you can use this property only in an event procedure or in a macro specified by the **OnPrint** event property setting.

You can set the default for this property by using the default control style or the **[DefaultControl](access.form.defaultcontrol.md)** property in Visual Basic.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
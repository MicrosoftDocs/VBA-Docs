---
title: Report.Cycle property (Access)
keywords: vbaac10.chm13821
f1_keywords:
- vbaac10.chm13821
ms.prod: access
api_name:
- Access.Report.Cycle
ms.assetid: 031194ca-f058-3a73-3551-f67d4e9bc27a
ms.date: 03/15/2019
localization_priority: Normal
---


# Report.Cycle property (Access)

You can use the **Cycle** property to specify what happens when you press the Tab key and the focus is in the last control on a report. Read/write **Byte**.


## Syntax

_expression_.**Cycle**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

The **Cycle** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|All Records|0|(Default) Pressing the Tab key from the last control on a report moves the focus to the first control in the tab order in the next record.|
|Current Record|1|Pressing the Tab key from the last control on a record moves the focus to the first control in the tab order in the same record.|
|Current Page|2|Pressing the Tab key from the last control on a page moves the focus back to the first control in the tab order on the page.|

You can set the **Cycle** property by using the report's property sheet, a macro, or Visual Basic.

You can set the **Cycle** property in any view.

When you press the Tab key on a report, the focus moves through the controls on the report according to each control's place in the tab order.

The **Cycle** property only controls the Tab key behavior on the report where the property is set. If a subreport control is in the tab order, after the subreport control receives the focus, the **Cycle** property setting for the subreport determines what happens when you press the Tab key.

To move the focus outside a subreport control, press Ctrl+Tab.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
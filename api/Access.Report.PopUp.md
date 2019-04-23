---
title: Report.PopUp property (Access)
keywords: vbaac10.chm13798
f1_keywords:
- vbaac10.chm13798
ms.prod: access
api_name:
- Access.Report.PopUp
ms.assetid: 76e82181-c5d5-01b2-c7ce-b2c78f237a75
ms.date: 03/20/2019
localization_priority: Normal
---


# Report.PopUp property (Access)

Specifies whether a report opens as a pop-up window. Read/write **Boolean**.


## Syntax

_expression_.**PopUp**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

The **PopUp** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Yes|**True**|The report opens as a pop-up window. It remains on top of all other Microsoft Access windows.|
|No|**False**|(Default) The report isn't a pop-up window.|

The **PopUp** property can be set only in Design view.

To specify the type of border that you want on a pop-up window, use the **BorderStyle** property. You typically set the **BorderStyle** property to Thin for pop-up windows.

Setting the **PopUp** property to Yes makes the report a pop-up window only when you do one of the following:

- Open it in Form view from the Database window.    
- Open it in Form view by using a macro or Visual Basic.    
- Switch from Design view to Form view.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
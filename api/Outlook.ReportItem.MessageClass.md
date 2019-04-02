---
title: ReportItem.MessageClass property (Outlook)
keywords: vbaol11.chm1653
f1_keywords:
- vbaol11.chm1653
ms.prod: outlook
api_name:
- Outlook.ReportItem.MessageClass
ms.assetid: 096bfebc-20eb-ea36-cff8-a96a514b5903
ms.date: 06/08/2017
localization_priority: Normal
---


# ReportItem.MessageClass property (Outlook)

Returns or sets a  **String** representing the message class for the Outlook item. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [ReportItem](Outlook.ReportItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[ReportItem Object](Outlook.ReportItem.md)



[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
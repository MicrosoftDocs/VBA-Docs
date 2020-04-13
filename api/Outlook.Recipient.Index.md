---
title: Recipient.Index property (Outlook)
keywords: vbaol11.chm2349
f1_keywords:
- vbaol11.chm2349
ms.prod: outlook
api_name:
- Outlook.Recipient.Index
ms.assetid: fe2ef09a-0046-1f82-e2ad-2e4cbb5a403f
ms.date: 06/08/2017
localization_priority: Normal
---


# Recipient.Index property (Outlook)

Returns a **Long** indicating the position of the object within the collection. Read-only.


## Syntax

_expression_.**Index**

_expression_ A variable that represents a [Recipient](Outlook.Recipient.md) object.


## Remarks

The **Index** property is only valid during the current session and can change as objects are added to and deleted from the collection. The first object in the collection has an **Index** value of 1.


## See also


[Recipient Object](Outlook.Recipient.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
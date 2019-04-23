---
title: Attachment.Index property (Outlook)
keywords: vbaol11.chm2367
f1_keywords:
- vbaol11.chm2367
ms.prod: outlook
api_name:
- Outlook.Attachment.Index
ms.assetid: 639ebc08-40a1-12ab-d9e1-6754add14b24
ms.date: 06/08/2017
localization_priority: Normal
---


# Attachment.Index property (Outlook)

Returns a  **Long** indicating the position of the object within the collection. Read-only.


## Syntax

_expression_.**Index**

_expression_ A variable that represents an [Attachment](Outlook.Attachment.md) object.


## Remarks

The  **Index** property is only valid during the current session and can change as objects are added to and deleted from the collection. The first object in the collection has an **Index** value of 1.


## See also


[Attachment Object](Outlook.Attachment.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
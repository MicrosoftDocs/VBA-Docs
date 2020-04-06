---
title: Conversation.Parent property (Outlook)
keywords: vbaol11.chm3385
f1_keywords:
- vbaol11.chm3385
ms.prod: outlook
api_name:
- Outlook.Conversation.Parent
ms.assetid: e1b3f294-227a-27d9-84db-042da1be0caa
ms.date: 06/08/2017
localization_priority: Normal
---


# Conversation.Parent property (Outlook)

Returns the parent  **Object** of the specified **[Conversation](Outlook.Conversation.md)** object. Read-only.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a [Conversation](Outlook.Conversation.md) object.


## Remarks

The parent of a  **Conversation** object is the first item in the conversation.

 If all items in the conversation are deleted after the **Conversation** object has been obtained, the **Parent** property returns **Null** (**Nothing** in Visual Basic).


## See also


[Conversation Object](Outlook.Conversation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
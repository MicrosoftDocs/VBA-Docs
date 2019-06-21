---
title: ServerPublishOptions.Stat property (Visio)
keywords: vis_sdr.chm17914420
f1_keywords:
- vis_sdr.chm17914420
ms.prod: visio
api_name:
- Visio.ServerPublishOptions.Stat
ms.assetid: 2a9c3a1a-ece6-9fd5-d470-eee7f9db8c57
ms.date: 06/08/2017
localization_priority: Normal
---


# ServerPublishOptions.Stat property (Visio)

Returns status information for an object. Read-only.


## Syntax

_expression_.**Stat**

_expression_ A variable that represents a **[ServerPublishOptions](Visio.ServerPublishOptions.md)** object.


## Return value

 **[VisStatCodes](Visio.visstatcodes.md)**


## Remarks

If an object is a reference to an entity in a document, and if that document closes, the  **Stat** property returns a value in which the **visStatClosed** bit is set.

If an object is a reference to an entity that has been deleted, the  **Stat** property returns a value in which the **visStatDeleted** bit is set.

A Component Object Model (COM) object, such as a Microsoft Visio  **[Document](Visio.Document.md)** object, lives as long as it is held (pointed to) by a client, even if the object is logically in a deleted or closed state.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: Validation.Stat property (Visio)
keywords: vis_sdr.chm18014420
f1_keywords:
- vis_sdr.chm18014420
ms.prod: visio
api_name:
- Visio.Validation.Stat
ms.assetid: abf46f37-4e0f-39c7-9368-6201b4bd5cf4
ms.date: 06/08/2017
localization_priority: Normal
---


# Validation.Stat property (Visio)

Returns status information for an object. Read-only.


## Syntax

_expression_.**Stat**

_expression_ A variable that represents a **[Validation](Visio.Validation.md)** object.


## Return value

 **[VisStatCodes](Visio.visstatcodes.md)**


## Remarks

If an object is a reference to an entity in a document, and if that document closes, the  **Stat** property returns a value in which the **visStatClosed** bit is set.

If an object is a reference to an entity that has been deleted, the  **Stat** property returns a value in which the **visStatDeleted** bit is set.

A Component Object Model (COM) object, such as a Microsoft Visio  **[Document](Visio.Document.md)** object, lives as long as it is held (pointed to) by a client, even if the object is logically in a deleted or closed state.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
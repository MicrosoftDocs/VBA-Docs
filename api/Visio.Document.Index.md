---
title: Document.Index Property (Visio)
keywords: vis_sdr.chm10513695
f1_keywords:
- vis_sdr.chm10513695
ms.prod: visio
api_name:
- Visio.Document.Index
ms.assetid: f72e68b9-c249-b4df-14ae-669509100546
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Index Property (Visio)

Gets the ordinal position of a  **Document** object in the **Documents** collection. Read-only.


## Syntax

 _expression_. `Index`

 _expression_ A variable that represents a [Document](./Visio.Document.md) object.


## Return value

Integer


## Remarks

Most collections are indexed starting with 1 rather than zero (0), so the index of the first element is 1, the index of the second element is 2, and so forth. The index of the last element in a collection is the same as the value of that collection's  **Count** property. You can iterate through a collection by using these index values. Adding objects to or deleting objects from a collection can change the index values of other objects in the collection.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
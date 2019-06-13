---
title: ReaderSpread.Height property (Publisher)
keywords: vbapb10.chm524296
f1_keywords:
- vbapb10.chm524296
ms.prod: publisher
api_name:
- Publisher.ReaderSpread.Height
ms.assetid: dfb84798-da3f-516b-22cd-0ba2a63ff39d
ms.date: 06/13/2019
localization_priority: Normal
---


# ReaderSpread.Height property (Publisher)

Returns a **Single** that represents the height, in [points](../language/glossary/vbe-glossary.md#point), of the page. Read-only.


## Syntax

_expression_.**Height**

_expression_ A variable that represents a **[ReaderSpread](Publisher.ReaderSpread.md)** object.


## Remarks

The valid range for the **Height** property depends on the size of the application workspace and the position of the object within the workspace. 

For centered objects on non-banner page sizes, the **Height** property may be 0.0 to 50.0 inches. For centered objects on banner page sizes, the **Height** property may be 0.0 to 241.0 inches.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
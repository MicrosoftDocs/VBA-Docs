---
title: LayoutGuides.VerticalBaseLineOffset property (Publisher)
keywords: vbapb10.chm1114133
f1_keywords:
- vbapb10.chm1114133
ms.prod: publisher
api_name:
- Publisher.LayoutGuides.VerticalBaseLineOffset
ms.assetid: 9a2f031c-4469-ca26-3e79-dfa556762e05
ms.date: 06/08/2019
localization_priority: Normal
---


# LayoutGuides.VerticalBaseLineOffset property (Publisher)

Returns a **Single** that represents the vertical baseline offset of the specified **LayoutGuides** object. Read/write.


## Syntax

_expression_.**VerticalBaseLineOffset**

_expression_ A variable that represents a **[LayoutGuides](Publisher.LayoutGuides.md)** object.


## Return value

Single


## Remarks

When setting the layout guide properties of a **Page** object, it must be returned from the **MasterPages** collection.


## Example

This example sets the vertical baseline offset of the **LayoutGuides** object to 12 for the second master page in the active document.

```vb
Dim objLayout As LayoutGuides 
Set objLayout = ActiveDocument.MasterPages(2).LayoutGuides 
objLayout.VerticalBaseLineOffset = 12 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
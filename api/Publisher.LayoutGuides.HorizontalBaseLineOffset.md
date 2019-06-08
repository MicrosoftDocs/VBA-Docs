---
title: LayoutGuides.HorizontalBaseLineOffset property (Publisher)
keywords: vbapb10.chm1114131
f1_keywords:
- vbapb10.chm1114131
ms.prod: publisher
api_name:
- Publisher.LayoutGuides.HorizontalBaseLineOffset
ms.assetid: b80d2114-8132-db13-a50d-ce904dbe5919
ms.date: 06/08/2019
localization_priority: Normal
---


# LayoutGuides.HorizontalBaseLineOffset property (Publisher)

Returns a **Single** that represents the horizontal baseline offset of the specified **LayoutGuides** object. Read/write.


## Syntax

_expression_.**HorizontalBaseLineOffset**

_expression_ A variable that represents a **[LayoutGuides](Publisher.LayoutGuides.md)** object.


## Return value

Single


## Remarks

When setting the layout guide properties of a **Page** object, it must be returned from the **MasterPages** collection.


## Example

This example sets the horizontal baseline offset of the **LayoutGuides** object to 12 for the second master page in the active document.

```vb
Dim objLayout As LayoutGuides 
Set objLayout = ActiveDocument.MasterPages(2).LayoutGuides 
objLayout.HorizontalBaseLineSpacing = 12 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
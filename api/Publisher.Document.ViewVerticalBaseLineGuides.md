---
title: Document.ViewVerticalBaseLineGuides property (Publisher)
keywords: vbapb10.chm196729
f1_keywords:
- vbapb10.chm196729
ms.prod: publisher
api_name:
- Publisher.Document.ViewVerticalBaseLineGuides
ms.assetid: 711335ab-237b-65a2-534a-7635cfba474e
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.ViewVerticalBaseLineGuides property (Publisher)

Sets or returns a **Boolean** that represents whether or not the vertical baseline guides are visible in the specified **Document** object. **True** if they are visible. **False** if they are not visible. Read/write.


## Syntax

_expression_.**ViewVerticalBaseLineGuides**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

Boolean


## Remarks

The default setting for this property is **False**.


## Example

The following example makes the vertical baseline guides visible in the active document.

```vb
Dim objDocument As Document 
Set objDocument = ActiveDocument 
objDocument.ViewVerticalBaseLineGuides = True 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
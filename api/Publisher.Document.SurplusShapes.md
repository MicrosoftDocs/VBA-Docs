---
title: Document.SurplusShapes property (Publisher)
keywords: vbapb10.chm196754
f1_keywords:
- vbapb10.chm196754
ms.prod: publisher
api_name:
- Publisher.Document.SurplusShapes
ms.assetid: 8c1c5fee-bea0-1660-a4a5-b465879d6ec9
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.SurplusShapes property (Publisher)

Returns a **[ShapeRange](publisher.shaperange.md)** object that represents the collection of surplus shapes that Microsoft Publisher places under **Extra Content** in the **Format Publication** task pane after the document template (wizard) is changed by using the **[ChangeDocument](Publisher.Document.ChangeDocument.md)** method or by using the **Change Template** command in the user interface. Read-only.


## Syntax

_expression_.**SurplusShapes**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

ShapeRange


## Remarks

Publisher classifies a shape as surplus if it does not fit neatly into the new template after the template is changed.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
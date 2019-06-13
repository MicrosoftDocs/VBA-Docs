---
title: Shape.IsExcess property (Publisher)
keywords: vbapb10.chm2228377
f1_keywords:
- vbapb10.chm2228377
ms.prod: publisher
api_name:
- Publisher.Shape.IsExcess
ms.assetid: 217689d6-7508-92ab-3828-e61fc70f0993
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.IsExcess property (Publisher)

Indicates whether the parent **Shape** object is an excess shape after the document template (wizard) is changed by using the **[Document.ChangeDocument](Publisher.Document.ChangeDocument.md)** method or by using the **Change Template** command in the user interface. Microsoft Publisher places any excess shape under **Extra Content** in the **Format Publication** task pane. Read-only.


## Syntax

_expression_.**IsExcess**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Return value

**[MsoTriState](office.msotristate.md)**


## Remarks

Publisher classifies a shape as excess (surplus) if that shape does not fit neatly into the new template after the template is changed.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
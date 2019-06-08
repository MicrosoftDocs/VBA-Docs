---
title: LineFormat.GradientStyle property (Publisher)
keywords: vbapb10.chm3408151
f1_keywords:
- vbapb10.chm3408151
ms.prod: publisher
ms.assetid: e5416db9-a145-8f71-2d75-1720191922bb
ms.date: 06/08/2019
localization_priority: Normal
---


# LineFormat.GradientStyle property (Publisher)

Returns the gradient style for the specified line. Read-only **[MsoGradientStyle](office.msogradientstyle.md)**.


## Syntax

_expression_.**GradientStyle**

_expression_ A variable that represents a **[LineFormat](Publisher.LineFormat.md)** object.


## Return value

MsoGradientStyle


## Remarks

Attempting to return this property for a line that doesn't have a gradient generates an error. Use the **[Type](Publisher.lineformat.type.md)** property to determine whether the line has a gradient.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: Presentation.Windows property (PowerPoint)
keywords: vbapp10.chm583017
f1_keywords:
- vbapp10.chm583017
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Windows
ms.assetid: ce04c680-ef68-5014-ce78-0d48d1f3b9e6
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.Windows property (PowerPoint)

Returns a **[DocumentWindows](PowerPoint.DocumentWindows.md)** collection that represents all document windows associated with the specified presentation. Read-only.


## Syntax

_expression_.**Windows**

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

DocumentWindows


## Remarks

This property doesn't return any slide show windows associated with the presentation.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.PowerPoint** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.PowerPoint._Presentation.Windows**
    

## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
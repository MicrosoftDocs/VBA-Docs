---
title: Report.MouseWheel property (Access)
keywords: vbaac10.chm13872
f1_keywords:
- vbaac10.chm13872
ms.prod: access
api_name:
- Access.Report.MouseWheel
ms.assetid: ea9d6443-abfd-6140-e167-548f4aafd342
ms.date: 03/09/2019
localization_priority: Normal
---


# Report.MouseWheel property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[MouseWheel](access.report.mousewheel(even).md)** event occurs. Read/write.


## Syntax

_expression_.**MouseWheel**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **MouseWheel** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
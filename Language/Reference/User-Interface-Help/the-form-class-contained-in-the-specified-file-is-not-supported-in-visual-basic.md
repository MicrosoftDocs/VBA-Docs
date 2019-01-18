---
title: The form class contained in the specified file is not supported in Visual Basic for Applications; the file can't be loaded.
keywords: vblr6.chm50053
f1_keywords:
- vblr6.chm50053
ms.prod: office
ms.assetid: 9b6dc45c-2076-3e78-4bec-e6d5b913d282
ms.date: 06/08/2017
localization_priority: Normal
---


# The form class contained in the specified file is not supported in Visual Basic for Applications; the file can't be loaded.

Visual Basic for Applications only supports  **[UserForm](userform-window.md)** form objects. This error has the following cause and solution:



- You are trying to load a form from an earlier version of Visual Basic or a form from a current standalone version of Visual Basic. It is probably easiest to redesign the form using the  **[UserForm](userform-window.md)** tools within Visual Basic for Applications. If you can save the form in [ASCII](../../Glossary/vbe-glossary.md#ascii-character-set) format, you may be able to modify the form and import it. However, you must completely understand the formats and limitations of the various form packages.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
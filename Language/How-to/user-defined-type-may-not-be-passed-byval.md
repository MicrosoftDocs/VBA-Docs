---
title: User-defined type may not be passed ByVal (VBA)
keywords: vblr6.chm1040140
f1_keywords:
- vblr6.chm1040140
ms.prod: office
ms.assetid: 1fbfeef6-b92d-03ca-aeec-4cf4c0d8d972
ms.date: 12/27/2018
localization_priority: Normal
---


# User-defined type may not be passed ByVal

[User-defined types](../Glossary/vbe-glossary.md#user-defined-type) can only be passed [by reference](../Glossary/vbe-glossary.md#by-reference) (the default), not [by value](../Glossary/vbe-glossary.md#by-value). The error may not be reported until the call is made. This error has the following cause and solution:

You placed a **ByVal** keyword in the definition of a [parameter](../Glossary/vbe-glossary.md#parameter) that represented a user-defined type.
    
Remove the **ByVal** keyword. To keep changes from being propagated back to the caller, **Dim** a temporary [variable](../Glossary/vbe-glossary.md#variable) of the type and pass the temporary variable into the [procedure](../Glossary/vbe-glossary.md#procedure).
    
For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

## See also

- [Visual Basic how-to topics](../reference/user-interface-help/visual-basic-how-to-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
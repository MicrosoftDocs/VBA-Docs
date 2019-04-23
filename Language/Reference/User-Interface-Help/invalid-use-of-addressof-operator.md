---
title: Invalid use of AddressOf operator
keywords: vblr6.chm1107785
f1_keywords:
- vblr6.chm1107785
ms.prod: office
ms.assetid: 96ce20a6-175e-a006-f0fe-98080d630c7f
ms.date: 11/19/2018
localization_priority: Normal
---


# Invalid use of AddressOf operator

The **[AddressOf operator](addressof-operator.md)** modifies an [argument](../../Glossary/vbe-glossary.md#argument) to pass the address of a function rather than passing the result of the function call. This error has the following cause and solution:

- You tried to use **AddressOf** with the name of a class method. Only the names of Visual Basic procedures in a .bas module can be modified with **AddressOf**. You can't specify a class method.
    
- The procedure name modified by **AddressOf** is defined in a [module](../../Glossary/vbe-glossary.md#module) in a different [project](../../Glossary/vbe-glossary.md#project).
    
- You tried to modify the name of a DLL function or a function defined in a [type library](../../Glossary/vbe-glossary.md#type-library) with **AddressOf**.
    
- DLL and type library functions can't be modified with **AddressOf**. The procedure definition must be in a module in the current project. Move the definition to a module in this project or include its current module in the project.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

## See also

- [Operator summary](operator-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
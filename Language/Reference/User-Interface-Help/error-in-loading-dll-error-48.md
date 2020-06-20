---
title: Error in loading DLL (Error 48)
keywords: vblr6.chm1011129
f1_keywords:
- vblr6.chm1011129
ms.prod: office
ms.assetid: 1dc4647e-3a73-9873-b10f-76b6c6ef1092
ms.date: 06/08/2017
localization_priority: Normal
---


# Error in loading DLL (Error 48)

A [dynamic link library (DLL)](../../Glossary/vbe-glossary.md#dynamic-link-library-dll) is a library specified in the **Lib** clause of a **Declare** statement. This error has the following causes and solutions:



- The file isn't DLL-executable. If the file is a source-text file, it must be compiled and linked to DLL executable form.
    
- The file isn't a Microsoft Windows DLL. Obtain the Microsoft Windows DLL equivalent of the file.
    
- The file is an early Microsoft Windows DLL that is incompatible with Microsoft Windows protect mode. Obtain an updated version of the DLL.
    
- The DLL references another DLL that isn't present. Obtain the referenced DLL and make it available to the other DLL.
    
- The DLL or one of the referenced DLLs isn't in a directory specified by your path. Move the DLL to a referenced directory or place its current directory on the path.

- The DLL is not the same bit version as the operating environment of VBA. For more information about bit version incompatibility, see [Mixing 32 and 64 bit DLLs](https://software.intel.com/en-us/forums/intel-fortran-compiler/topic/515978) on the Intel developer forum.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

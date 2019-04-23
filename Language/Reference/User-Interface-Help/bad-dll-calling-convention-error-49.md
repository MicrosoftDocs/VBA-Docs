---
title: Bad DLL calling convention (Error 49)
keywords: vblr6.chm1000049
f1_keywords:
- vblr6.chm1000049
ms.prod: office
ms.assetid: bf8ccbbe-87fa-8de4-dfd6-97642feee087
ms.date: 06/08/2017
localization_priority: Normal
---


# Bad DLL calling convention (Error 49)

[Arguments](../../Glossary/vbe-glossary.md#argument) passed to a [dynamic-link library](../../Glossary/vbe-glossary.md#dynamic-link-library-dll) (DLL) or Macintosh code resource routine must exactly match those expected by the routine. Calling conventions deal with number, type, and order of arguments. This error has the following causes and solutions:



- Your program is calling a routine in a DLL (in Windows) or a code resource (on the Macintosh) that's being passed the wrong type of arguments. Make sure all argument types agree with those specified in the declaration of the routine you are calling.
    
- Your program is calling a routine in a DLL (in Windows) or a code resource (on the Macintosh) that's being passed the wrong number of arguments. Make sure you are passing the same number of arguments indicated in the declaration of the routine you are calling.
    
- Your program is calling a routine in a DLL, but isn't using the StdCall calling convention. If the DLL routine expects arguments [by value](../../Glossary/vbe-glossary.md#by-value), then make sure  **ByVal** is specified for those arguments in the declaration for the routine.
    
- Your  **Declare** statement for a Windows DLL includes **CDecl**. The **CDecl** keyword applies only to the Macintosh.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
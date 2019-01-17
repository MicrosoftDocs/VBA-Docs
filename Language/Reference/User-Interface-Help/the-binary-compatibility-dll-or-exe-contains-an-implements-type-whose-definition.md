---
title: The binary compatibility DLL or EXE contains an Implements type whose definition cannot be found
keywords: vblr6.chm1040373
f1_keywords:
- vblr6.chm1040373
ms.prod: office
ms.assetid: 4cace415-821b-d0d6-64ca-ccc4fe207f4a
ms.date: 06/08/2017
localization_priority: Normal
---


# The binary compatibility DLL or EXE contains an Implements type whose definition cannot be found

If you have a Binary Compatible server which implements an interface that is contained in another DLL, you must be careful when recompiling it. This warning has the following cause and solution:



- The other DLL was recompiled as Project Compatible which changes the interface GUID. Since this is not a visible change, this can be an unexpected error. This can also occur if someone gives you a Project Compatible DLL to reference. Basically, this error occurs when a project's binary compatible DLL or EXE has a typelib with a broken reference. Broken references can occur when a referenced typelib is overwritten by another file (such as a re-compiled DLL/EXE), when you delete the typelib file, or when you move a referencing typelib over to a machine, but either don't move the referenced typelib or don't register the referenced typelib. One possible fix is to obtain a copy of the referenced typelib onto your machine and register it. You won't be able to use the old one because it was overwritten on recompile. Failing this, all that can be done is to stop using the DLL/EXE as your binary compatible version.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
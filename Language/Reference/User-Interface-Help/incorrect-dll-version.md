---
title: Incorrect DLL version
keywords: vblr6.chm1011194
f1_keywords:
- vblr6.chm1011194
ms.prod: office
ms.assetid: 6ff7118c-1764-8098-9728-10146e270312
ms.date: 06/08/2017
localization_priority: Normal
---


# Incorrect DLL version

Each version of Visual Basic works only with its corresponding [dynamic-link library (DLL)](../../Glossary/vbe-glossary.md#dynamic-link-library-dll) (Windows) or code resource (Macintosh). This error has the following cause and solution:



- Your version of the Visual Basic dynamic-link library or code resource doesn't match the version expected by this [host application](../../Glossary/vbe-glossary.md#host-application). The program is attempting to call routines in a DLL or code resource, but the version of the library or resource is inconsistent with either Visual Basic or the host application.
    
    Obtain the correct version of the library or resource, and make sure earlier versions don't precede the proper one on your path.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
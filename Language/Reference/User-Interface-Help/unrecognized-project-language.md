---
title: Unrecognized project language
keywords: vblr6.chm1032814
f1_keywords:
- vblr6.chm1032814
ms.prod: office
ms.assetid: deaf7459-f91f-2ad7-fb94-e954939a8b99
ms.date: 08/24/2018
localization_priority: Normal
---


# Unrecognized project language

The specified code [locale](../../Glossary/vbe-glossary.md#locale) for the [project](../../Glossary/vbe-glossary.md#project) to be loaded isn't currently supported by this application. This error has the following causes and solutions:

- The project was created on a system that supports the code locale, but was then moved to a system where that code locale isn't recognized. For example, the ole2nls.dll on the current machine may be a version that doesn't recognize the code locale. Install the proper [dynamic-link library (DLL)](../../Glossary/vbe-glossary.md#dynamic-link-library-dll) on the current system.
    
- The correct [object library](../../Glossary/vbe-glossary.md#object-library) for the project was not found.

  Make sure the correct object libraries are available, for example, make sure your path includes their directories.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

> [!TIP] 
> 2018-08-16: One working solution for this problem is to unselect the UTF-8 beta setting in the **Control Panel** > **Regional settings** > **Administrative settings**. The [original solution is in Vietnamese](https://blog.hocexcel.online/sua-loi-unrecognized-project-language-trong-vba.html) (use Google Translate to read).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

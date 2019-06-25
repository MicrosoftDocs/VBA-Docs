---
title: InvisibleApp.LanguageSettings property (Visio)
keywords: vis_sdr.chm17560035
f1_keywords:
- vis_sdr.chm17560035
ms.prod: visio
api_name:
- Visio.InvisibleApp.LanguageSettings
ms.assetid: 0aff05cd-7655-0671-9c43-e45988c5a172
ms.date: 06/26/2019
localization_priority: Normal
---


# InvisibleApp.LanguageSettings property (Visio)

Returns a reference to the Microsoft Office (MSO) **[LanguageSettings](office.languagesettings.md)** interface. Read-only.


## Syntax

_expression_.**LanguageSettings**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Return value

Object


## Remarks

After you use the **LanguageSettings** property to get a reference to the MSO **LanguageSettings** interface, you can use methods of that interface to get the locale identifier (LCID) for the language used when Office was installed, the user interface (UI) language, and the language for Help, as well as the current setting for the preferred language for editing in the UI.

However, you cannot use the **LanguageSettings** interface to change language settings; you can change language settings only in the **Microsoft Office Language Settings** dialog box (**Start** > **All Programs** > **Microsoft Office** > **Microsoft Office Tools** > **Microsoft Office Language Settings**).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
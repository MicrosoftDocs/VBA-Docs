---
title: GetAllSettings function (Visual Basic for Applications)
keywords: vblr6.chm1020903
f1_keywords:
- vblr6.chm1020903
ms.prod: office
ms.assetid: f87675b2-d14e-593d-94ab-259ab8da094d
ms.date: 12/12/2018
localization_priority: Normal
---


# GetAllSettings function

Returns a list of key settings and their respective values (originally created with **[SaveSetting](savesetting-statement.md)**) from an application's entry in the Windows [registry](../../Glossary/vbe-glossary.md#registry) or (on the Macintosh) information in the application's initialization file.

## Syntax

**GetAllSettings**(_appname_, _section_)

<br/>

The **GetAllSettings** function syntax has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_appname_|Required. [String expression](../../Glossary/vbe-glossary.md#string-expression) containing the name of the application or [project](../../Glossary/vbe-glossary.md#project) whose key settings are requested. On the Macintosh, this is the filename of the initialization file in the Preferences folder in the System folder.|
|_section_|Required. String **expression** containing the name of the section whose key settings are requested. **GetAllSettings** returns a [Variant](../../Glossary/vbe-glossary.md#variant-data-type) whose contents is a two-dimensional [array](../../Glossary/vbe-glossary.md#array) of strings containing all the key settings in the specified section and their corresponding values.|

## Remarks

**GetAllSettings** returns an uninitialized **Variant** if either _appname_ or _section_ does not exist.

## Example

This example first uses the **SaveSetting** statement to make entries in the Windows registry for the application specified as _appname_, and then uses the **GetAllSettings** function to display the settings. Note that application names and _section_ names can't be retrieved with **GetAllSettings**. Finally, the **[DeleteSetting](deletesetting-statement.md)** statement removes the application's entries.


```vb
' Variant to hold 2-dimensional array returned by GetAllSettings
' Integer to hold counter.
Dim MySettings As Variant, intSettings As Integer
' Place some settings in the registry.
SaveSetting appname := "MyApp", section := "Startup", _
key := "Top", setting := 75
SaveSetting "MyApp","Startup", "Left", 50
' Retrieve the settings.
MySettings = GetAllSettings(appname := "MyApp", section := "Startup")
    For intSettings = LBound(MySettings, 1) To UBound(MySettings, 1)
        Debug.Print MySettings(intSettings, 0), MySettings(intSettings, 1)
    Next intSettings
DeleteSetting "MyApp", "Startup"

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
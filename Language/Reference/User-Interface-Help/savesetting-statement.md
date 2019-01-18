---
title: SaveSetting statement (VBA)
keywords: vblr6.chm1020904
f1_keywords:
- vblr6.chm1020904
ms.prod: office
ms.assetid: f15549da-3c84-0991-592a-9d715fd488f3
ms.date: 12/10/2018
localization_priority: Normal
---


# SaveSetting statement

Saves or creates an application entry in the application's entry in the Windows [registry](../../Glossary/vbe-glossary.md#registry) or (on the Macintosh) information in the application's initialization file.

## Syntax

**SaveSetting** _appname_, _section_, _key_, _setting_

<br/>

The **SaveSetting** statement syntax has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_appname_|Required. [String expression](../../Glossary/vbe-glossary.md#string-expression) containing the name of the application or [project](../../Glossary/vbe-glossary.md#project) to which the setting applies. On the Macintosh, this is the filename of the initialization file in the Preferences folder in the System folder.|
|_section_|Required. String expression containing the name of the section where the key setting is being saved.|
|_key_|Required. String expression containing the name of the key setting being saved.|
|_setting_|Required. [Expression](../../Glossary/vbe-glossary.md#expression) containing the value that _key_ is being set to.|

## Remarks

An error occurs if the key setting can't be saved for any reason.

The root of these registry settings is: `Computer\HKEY_CURRENT_USER\Software\VB and VBA Program Settings`.

## Example

The following example first uses the **SaveSetting** statement to make entries in the Windows registry (or .ini file on 16-bit Windows platforms) for the application, and then uses the **[DeleteSetting](deletesetting-statement.md)** statement to remove them.

```vb
' Place some settings in the registry. 
SaveSetting appname := "MyApp", section := "Startup", _ 
 key := "Top", setting := 75 
SaveSetting "MyApp","Startup", "Left", 50 
' Remove section and all its settings from registry. 
DeleteSetting "MyApp", "Startup" 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: ThemeColorScheme.Save method (Office)
ms.prod: office
api_name:
- Office.ThemeColorScheme.Save
ms.assetid: 5ca73773-583b-dbf4-6bde-bc6fa26c66a2
ms.date: 01/25/2019
localization_priority: Normal
---


# ThemeColorScheme.Save method (Office)

Saves the color scheme of a Microsoft Office theme to a file.


## Syntax

_expression_.**Save** (_FileName_)

_expression_ An expression that returns a **[ThemeColorScheme](Office.ThemeColorScheme.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the file.|

## Example

The following example saves the color scheme for an Office theme to a file.

```vb
ThemeColorScheme.Save("C:\myThemeColorScheme.xml") 

```


## See also

- [ThemeColorScheme object members](overview/Library-Reference/themecolorscheme-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
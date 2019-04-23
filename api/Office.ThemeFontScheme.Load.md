---
title: ThemeFontScheme.Load method (Office)
ms.prod: office
api_name:
- Office.ThemeFontScheme.Load
ms.assetid: a9ac928e-904f-70bd-1e96-932243204d73
ms.date: 01/29/2019
localization_priority: Normal
---


# ThemeFontScheme.Load method (Office)

Loads the font scheme of a Microsoft Office theme from a file.


## Syntax

_expression_.**Load** (_FileName_)

_expression_ An expression that returns a **[ThemeFontScheme](Office.ThemeFontScheme.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the font scheme file.|

## Example

The following example loads a theme font scheme from a file.


```vb
ThemeFontScheme.Load ("C:\myThemeFontScheme.xml")
```


## See also

- [ThemeFontScheme object members](overview/Library-Reference/themefontscheme-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
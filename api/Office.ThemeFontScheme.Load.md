---
title: ThemeFontScheme.Load Method (Office)
ms.prod: office
api_name:
- Office.ThemeFontScheme.Load
ms.assetid: a9ac928e-904f-70bd-1e96-932243204d73
ms.date: 06/08/2017
---


# ThemeFontScheme.Load Method (Office)

Loads the font scheme of a Microsoft Office theme from a file.


## Syntax

 _expression_. `Load`( `_FileName_` )

 _expression_ An expression that returns a [ThemeFontScheme](./Office.ThemeFontScheme.md) object.


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


[ThemeFontScheme Object](Office.ThemeFontScheme.md)



[ThemeFontScheme Object Members](./overview/Library-Reference/themefontscheme-members-office.md)


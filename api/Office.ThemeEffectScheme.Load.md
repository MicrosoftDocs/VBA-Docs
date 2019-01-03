---
title: ThemeEffectScheme.Load method (Office)
ms.prod: office
api_name:
- Office.ThemeEffectScheme.Load
ms.assetid: 9bf428f7-bda8-c6d7-1688-05466f242280
ms.date: 06/08/2017
---


# ThemeEffectScheme.Load method (Office)

Loads the effects scheme of a Microsoft Office theme from a file.


## Syntax

 _expression_. `Load`( `_FileName_` )

 _expression_ An expression that returns a [ThemeEffectScheme](Office.ThemeEffectScheme.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the effect scheme file.|

## Example

The following example loads a theme effect scheme from a file.


```vb
tesEffectScheme.Load("C:\myThemeEffectScheme.eftx") 

```


## See also


[ThemeEffectScheme Object](Office.ThemeEffectScheme.md)



[ThemeEffectScheme Object Members](./overview/Library-Reference/themeeffectscheme-members-office.md)


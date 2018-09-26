---
title: ThemeFontScheme.Save Method (Office)
ms.prod: office
api_name:
- Office.ThemeFontScheme.Save
ms.assetid: 4adbeac7-b5cf-327e-f999-4dd2d721755d
ms.date: 06/08/2017
---


# ThemeFontScheme.Save Method (Office)

Saves the font scheme of a Microsoft Office theme to a file.


## Syntax

 _expression_. `Save`( `_FileName_` )

 _expression_ An expression that returns a [ThemeFontScheme](./Office.ThemeFontScheme.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the file.|

## Example

The following example saves the Office theme font scheme to a file. 


```vb
ThemeFontScheme.Save("C:\myThemeFontScheme.xml")
```


## See also


[ThemeFontScheme Object](Office.ThemeFontScheme.md)



[ThemeFontScheme Object Members](./overview/Library-Reference/themefontscheme-members-office.md)


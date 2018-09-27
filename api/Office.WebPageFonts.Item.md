---
title: WebPageFonts.Item Property (Office)
keywords: vbaof11.chm225002
f1_keywords:
- vbaof11.chm225002
ms.prod: office
api_name:
- Office.WebPageFonts.Item
ms.assetid: 2f7f1286-749e-3598-8091-16c896bc4842
ms.date: 06/08/2017
---


# WebPageFonts.Item Property (Office)

Gets a  **WebPageFont** object from the **WebPageFonts** collection for a particular value of **MsoCharacterSet**. Read-only.


## Syntax

 _expression_. `Item`( `_Index_` )

 _expression_ Required. A variable that represents a '[WebPageFonts](Office.WebPageFonts.md)' object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**MsoCharacterSet**|The specified character set.|

## Example

The following example uses the  **Item** property to set "myFont" to the **WebPageFont** object for the **English/Western European/Other Latin Script** character set in the active application.


```vb
Dim myFont As WebPageFont 
Set myFont = _ 
 Application.DefaultWebOptions.Fonts. _ 
 Item(msoCharacterSetEnglishWesternEuropeanOtherLatinScript)
```


## See also


[WebPageFonts Object](Office.WebPageFonts.md)



[WebPageFonts Object Members](./overview/Library-Reference/webpagefonts-members-office.md)


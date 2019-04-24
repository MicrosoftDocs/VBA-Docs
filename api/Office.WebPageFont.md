---
title: WebPageFont object (Office)
keywords: vbaof11.chm224000
f1_keywords:
- vbaof11.chm224000
ms.prod: office
api_name:
- Office.WebPageFont
ms.assetid: daf3c079-520d-68bd-ec02-027776074505
ms.date: 01/29/2019
localization_priority: Normal
---


# WebPageFont object (Office)

Represents the default font used when documents are saved as webpages for a particular character set.


## Remarks

Use the **WebPageFont** object to describe the proportional font, proportional font size, fixed-width font, and fixed-width font size for any available character set.


## Example

The following example sets the proportional font and proportional font size for the **WebPageFont** object's "myFont".


```vb
With myFont 
 ProportionalFont = Verdana 
 ProportionalFontSize = 14
```


## See also

- [WebPageFont object members](overview/Library-Reference/webpagefont-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
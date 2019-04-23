---
title: Page.Height property (Word)
keywords: vbawd10.chm11075589
f1_keywords:
- vbawd10.chm11075589
ms.prod: word
api_name:
- Word.Page.Height
ms.assetid: fe097fed-868b-cb09-f2ad-d53cda76a426
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.Height property (Word)

Returns a  **Long** that represents the height of a page, in pixels.


## Syntax

_expression_.**Top**

 _expression_ An expression that represents a '[Page](Word.Page.md)' object.


## Remarks

The  **[Top](Word.Page.Top.md)** and **[Left](Word.Page.Left.md)** properties of the **Page** object always return 0 (zero) indicating the upper-left corner of the page. The **Height** and **[Width](Word.Page.Width.md)** properties return the height and width in points (72 points = 1 inch) of the paper size specified in the **Page Setup** dialog box or through the **[PageSetup](Word.PageSetup.md)** object. For example, for an 8-1/2 by 11 inch page in portrait mode, the **Height** property returns 792 and the **Width** property returns 612. All four of these properties are read-only.


## See also


[Page Object](Word.Page.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
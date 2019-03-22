---
title: Section.WillContinue property (Access)
keywords: vbaac10.chm12213
f1_keywords:
- vbaac10.chm12213
ms.prod: access
api_name:
- Access.Section.WillContinue
ms.assetid: e79785e6-87b8-dd9f-9659-341c2fd81bf5
ms.date: 03/23/2019
localization_priority: Normal
---


# Section.WillContinue property (Access)

Determines if the current section will continue on the following page. Read/write **Boolean**.


## Syntax

_expression_.**WillContinue**

_expression_ A variable that represents a **[Section](Access.Section.md)** object.


## Remarks

You can use this property to determine whether to show or hide certain controls, depending on the value of the property. For example, you may have a hidden label in a page header containing the text "Continued on next page." If the value of the **WillContinue** property is **True**, you can make the hidden label visible.


## Example

The following example displays a message box indicating whether the page header for the report **Product Summary** will continue on the following page.

```vb
MsgBox Reports("Product Summary").Section("PageHeaderSection").WillContinue
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
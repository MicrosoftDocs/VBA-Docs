---
title: IconView.XML property (Outlook)
keywords: vbaol11.chm2572
f1_keywords:
- vbaol11.chm2572
ms.prod: outlook
api_name:
- Outlook.IconView.XML
ms.assetid: d6876759-9266-77ab-c61e-e7d2eb240a96
ms.date: 06/08/2017
localization_priority: Normal
---


# IconView.XML property (Outlook)

Returns or sets a **String** value that specifies the XML definition of the view. Read/write.


## Syntax

_expression_.**XML**

_expression_ A variable that represents an [IconView](Outlook.IconView.md) object.


## Remarks

The XML definition describes the view type by using a series of tags and keywords corresponding to various properties of the view itself. When the view is created, the XML definition is parsed to render the settings for the new view.

To determine how the XML should be structured when creating views, you can create a view by using the Outlook user interface and then you can retrieve the  **XML** property for that view.


## See also


[IconView Object](Outlook.IconView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
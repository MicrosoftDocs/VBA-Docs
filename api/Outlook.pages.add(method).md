---
title: Pages.Add Method (Outlook Forms Script)
ms.prod: outlook
ms.assetid: be7bc499-8e25-440c-0ad9-2a6416ad8cea
ms.date: 06/08/2017
localization_priority: Normal
---


# Pages.Add Method (Outlook Forms Script)

Adds a  **[Page](Outlook.page.md)** to a **[Pages](Outlook.pages(object).md)** collection.


## Syntax

_expression_.**Add** (_bstrName_, _bstrCaption_, _lIndex_)

_expression_ A variable that represents a  **Pages** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|bstrName|Optional| **Variant**|Specifies the name of the object being added. If a name is not specified, the system generates a default name based on the rules of the application where the form is used.|
|bstrCaption|Optional| **Variant**|Specifies the caption to appear on a page. If a caption is not specified, the system generates a default caption based on the rules of the application where the form is used.|
|lIndex|Optional| **Variant**|Identifies the position of a page within a  **Pages** collection. If an index is not specified, the system appends the page to the end of the **Pages** collection and assigns the appropriate index value.|

## Return value

A  **Page** object that represents the added page.


## Remarks

The index value for the first  **Page** of a collection is 0, the value for the second **Page** is 1, and so on.

You can change the  **Name** property of the object at run time only if you added that control at run time with the **Add** method.


## See also


 [Pages Object](Outlook.pages(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
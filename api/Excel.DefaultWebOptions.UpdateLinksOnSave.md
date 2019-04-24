---
title: DefaultWebOptions.UpdateLinksOnSave property (Excel)
keywords: vbaxl10.chm660077
f1_keywords:
- vbaxl10.chm660077
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.UpdateLinksOnSave
ms.assetid: d2ae453f-8dc2-fe6c-a64c-574b22c781cd
ms.date: 04/25/2019
localization_priority: Normal
---


# DefaultWebOptions.UpdateLinksOnSave property (Excel)

**True** if hyperlinks and paths to all supporting files are automatically updated before you save the document as a webpage, ensuring that the links are up to date at the time the document is saved. **False** if the links are not updated. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**UpdateLinksOnSave**

_expression_ A variable that represents a **[DefaultWebOptions](Excel.DefaultWebOptions.md)** object.


## Remarks

You should set this property to **False** if the location where the document is saved is different from the final location on the web server, and the supporting files are not available at the first location.


## Example

This example specifies that links are not updated before the document is saved.

```vb
Application.DefaultWebOptions.UpdateLinksOnSave = False
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
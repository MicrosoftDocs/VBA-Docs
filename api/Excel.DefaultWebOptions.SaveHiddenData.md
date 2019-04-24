---
title: DefaultWebOptions.SaveHiddenData property (Excel)
keywords: vbaxl10.chm660074
f1_keywords:
- vbaxl10.chm660074
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.SaveHiddenData
ms.assetid: b1c09c39-3510-263c-3403-6e48d125279d
ms.date: 04/25/2019
localization_priority: Normal
---


# DefaultWebOptions.SaveHiddenData property (Excel)

**True** if data outside of the specified range is saved when you save the document as a webpage. This data may be necessary for maintaining formulas. **False** if data outside of the specified range is not saved with the webpage. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**SaveHiddenData**

_expression_ A variable that represents a **[DefaultWebOptions](Excel.DefaultWebOptions.md)** object.


## Remarks

If you choose not to save data outside of the specified range, references to that data in the formula are converted to static values. If the data is in another sheet or workbook, the result of the formula is saved as a static value.


## Example

This example prevents hidden data from being saved with webpages.

```vb
Application.DefaultWebOptions.SaveHiddenData = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: Form.FetchDefaults property (Access)
keywords: vbaac10.chm13555
f1_keywords:
- vbaac10.chm13555
ms.prod: access
api_name:
- Access.Form.FetchDefaults
ms.assetid: 3bbe8c57-e9ff-419a-d2b4-93cb966d6f30
ms.date: 03/12/2019
localization_priority: Normal
---


# Form.FetchDefaults property (Access)

Returns or sets a **Boolean** indicating whether Microsoft Access shows default values for new rows on the specified form before the row is saved. **True** if Access shows the default values for new rows on the specified form. Read/write.


## Syntax

_expression_.**FetchDefaults**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Example

The following example sets a form to show default values for new rows.

```vb
Forms(0).FetchDefaults = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
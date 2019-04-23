---
title: Form.NavigationButtons property (Access)
keywords: vbaac10.chm13365
f1_keywords:
- vbaac10.chm13365
ms.prod: access
api_name:
- Access.Form.NavigationButtons
ms.assetid: 23af1adc-67e9-b39d-772b-ddecf159f861
ms.date: 03/14/2019
localization_priority: Normal
---


# Form.NavigationButtons property (Access)

You can use the **NavigationButtons** property to specify whether navigation buttons and a record number box are displayed on a form. Read/write **Boolean**.


## Syntax

_expression_.**NavigationButtons**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

The default value is **True**.

Navigation buttons provide an efficient way to move to the first, previous, next, last, or blank (new) record. The record number box displays the number of the current record. The total number of records is displayed next to the navigation buttons. You can enter a number in the record number box to move to a particular record.

If you remove the navigation buttons from a form and want to create your owns means of navigation for the form, you can create custom navigation buttons and add them to the form.


## Example

The following example returns the value of the **NavigationButtons** property for the **Order Entry** form.

```vb
Dim b As Boolean 
b = Forms("Order Entry").NavigationButtons
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
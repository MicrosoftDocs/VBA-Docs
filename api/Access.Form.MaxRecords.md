---
title: Form.MaxRecords property (Access)
keywords: vbaac10.chm13484
f1_keywords:
- vbaac10.chm13484
ms.prod: access
api_name:
- Access.Form.MaxRecords
ms.assetid: 1c1ea306-7ab0-8818-2fb6-8ac377f73484
ms.date: 03/13/2019
localization_priority: Normal
---


# Form.MaxRecords property (Access)

Specifies the maximum number of records by a query or view. Read/write **Long**.


## Syntax

_expression_.**MaxRecords**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

When you set this property in Visual Basic, you use the ADO **MaxRecords** property.

Records are returned in the order specified by the query's ORDER BY clause.

You can use the **MaxRecords** property in situations where limited system resources might prohibit a large number of returned records.


## Example

To return the **MaxRecords** property for a form, you can use the following syntax.

```vb
Dim l As Longl = Forms(formname).MaxRecords
```

<br/>

To set the **MaxRecords** property, you can use the following syntax.

```vb
Forms(formname).MaxRecords = numrecords
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
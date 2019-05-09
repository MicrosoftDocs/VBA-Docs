---
title: Application.Workbooks property (Excel)
keywords: vbaxl10.chm183115
f1_keywords:
- vbaxl10.chm183115
ms.prod: excel
api_name:
- Excel.Application.Workbooks
ms.assetid: 5291a324-87d7-3916-ffee-34c3389cea13
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.Workbooks property (Excel)

Returns a **[Workbooks](Excel.Workbooks.md)** collection that represents all the open workbooks. Read-only.


## Syntax

_expression_.**Workbooks**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

Using this property without an object qualifier is equivalent to using Application.Workbooks.

The collection returned by the **Workbooks** property doesn't include open add-ins, which are a special kind of hidden workbook. You can, however, return a single open add-in if you know the file name. For example, `Workbooks("Oscar.xla")` returns the open add-in named "Oscar.xla" as a **Workbook** object.

> [!NOTE] 
> A workbook displayed in a Protected View window is not a member of the **Workbooks** collection. Instead, use the **[Workbook](Excel.ProtectedViewWindow.Workbook.md)** property of the **ProtectedViewWindow** object to access a workbook that is displayed in a Protected View window.


## Example

This example activates the workbook Book1.xls.

```vb
Workbooks("BOOK1").Activate
```

<br/>

This example opens the workbook Large.xls.

```vb
Workbooks.Open filename:="LARGE.XLS"
```

<br/>

This example saves changes to and closes all workbooks except the one that's running the example.

```vb
For Each w In Workbooks 
    If w.Name <> ThisWorkbook.Name Then 
        w.Close savechanges:=True 
    End If 
Next w
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

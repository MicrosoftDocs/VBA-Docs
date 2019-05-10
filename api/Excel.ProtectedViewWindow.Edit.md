---
title: ProtectedViewWindow.Edit method (Excel)
keywords: vbaxl10.chm914087
f1_keywords:
- vbaxl10.chm914087
ms.prod: excel
api_name:
- Excel.ProtectedViewWindow.Edit
ms.assetid: bdb626b2-ed4a-06d2-076c-5d242d23a162
ms.date: 05/09/2019
localization_priority: Normal
---


# ProtectedViewWindow.Edit method (Excel)

Opens the workbook that is open for editing in the specified Protected View window.


## Syntax

_expression_.**Edit** (_WriteResPassword_, _UpdateLinks_)

_expression_ A variable that represents a **[ProtectedViewWindow](Excel.ProtectedViewWindow.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _WriteResPassword_|Optional| **Variant**|The password required to write to a write-reserved workbook. If this argument is omitted and the workbook requires a password, the user will be prompted for the password.|
| _UpdateLinks_|Optional| **Variant**|Specifies the way that external references (links) in the file, such as the reference to a range in the Budget.xls workbook in the following formula `=SUM([Budget.xls]Annual!C10:C25)`, are updated.<br/><br/>If this argument is omitted, the user is prompted to specify how links will be updated. For more information about the values used by this parameter, see the Remarks section.<br/><br/>If Excel is opening a file in the WKS, WK1, or WK3 format and the  _UpdateLinks_ argument is 0, no charts are created; otherwise, Excel generates charts from the graphs attached to the file.|

## Return value

**[Workbook](Excel.Workbook.md)**


## Remarks

Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code.

You can specify one of the values, listed in the following table, in the _UpdateLinks_ parameter to determine whether external references (links) are updated when the workbook is opened.

|Value|Description|
|:-----|:-----|
|0|External references (links) will not be updated when the workbook is opened.|
|3|External references (links) will be updated when the workbook is opened.|


## Example

The following code example opens the workbook that is open in the active Protected View window for editing.

```vb
Dim pvWbk As Workbook 
 
Set pvWbk = ActiveProtectedViewWindow.Edit 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
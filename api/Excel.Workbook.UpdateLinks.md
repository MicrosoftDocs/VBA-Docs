---
title: Workbook.UpdateLinks property (Excel)
keywords: vbaxl10.chm199197
f1_keywords:
- vbaxl10.chm199197
ms.prod: excel
api_name:
- Excel.Workbook.UpdateLinks
ms.assetid: c8d374d7-0b32-eb32-fa29-ab496d6786e7
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.UpdateLinks property (Excel)

Returns or sets an **[XlUpdateLink](Excel.XlUpdateLinks.md)** constant indicating a workbook's setting for updating embedded OLE links. Read/write.


## Syntax

_expression_.**UpdateLinks**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

In this example, Microsoft Excel determines the setting for updating links and notifies the user.

```vb
Sub UseUpdateLinks() 
 
 Dim wkbOne As Workbook 
 
 Set wkbOne = Application.Workbooks(1) 
 
 Select Case wkbOne.UpdateLinks 
 Case xlUpdateLinksAlways 
 MsgBox "Links will always be updated " & _ 
 "for the specified workbook." 
 Case xlUpdateLinksNever 
 MsgBox "Links will never be updated " & _ 
 "for the specified workbook." 
 Case xlUpdateLinksUserSetting 
 MsgBox "Links will update according " & _ 
 "to user settting for the specified workbook." 
 End Select 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
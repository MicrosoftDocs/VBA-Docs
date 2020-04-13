---
title: PropertyPage.GetPageInfo method (Outlook)
keywords: vbaol11.chm381
f1_keywords:
- vbaol11.chm381
ms.prod: outlook
api_name:
- Outlook.PropertyPage.GetPageInfo
ms.assetid: 39243864-a81a-eaa6-965d-c1a5ac5ac781
ms.date: 06/08/2017
localization_priority: Normal
---


# PropertyPage.GetPageInfo method (Outlook)

Returns Help information about a custom property page.


## Syntax

_expression_. `GetPageInfo`( `_HelpFile_` , `_HelpContext_` )

_expression_ A variable that represents a [PropertyPage](Outlook.PropertyPage.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _HelpFile_|Required| **String**|Specifies the Help file associated with the property page.|
| _HelpContext_|Required| **Long**|Specifies the context ID of the Help topic associated with the property page.|

## Return value

An **HRESULT** value that represents the result of the method.


## Example

This Microsoft Visual Basic for Applications (VBA) example returns the name of the Help file and the context ID of the topic to be displayed.


```vb
Private Sub PropertyPage_GetPageInfo(HelpFile As String, HelpContext As Long) 
 HelpFile = "ProjPage.chm" 
 HelpContext = IDH_PageInfo 
End Sub
```


## See also


[PropertyPage Object](Outlook.PropertyPage.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
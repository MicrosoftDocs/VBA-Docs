---
title: Document.Protect method (Word)
keywords: vbawd10.chm158007763
f1_keywords:
- vbawd10.chm158007763
ms.prod: word
ms.assetid: 727bafe9-48ea-6b2f-2262-778f66487cbd
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Protect method (Word)

Protects the specified document from unauthorized changes.


## Syntax

_expression_.**Protect** (_Type_, _NoReset_, _Password_, _UseIRM_, _EnforceStyleLock_)

_expression_ A variable that represents a **[Document](./Word.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **WdProtectionType**|The type of protection to apply.|
| _NoReset_|Optional|**Variant**| **False** to reset form fields to their default values; **True** to retain the current form field values if the document is protected. If _Type_ is not **wdAllowOnlyFormFields**,  _NoReset_ is ignored.|
| _Password_|Optional|**Variant**|If supplied, the password to be able to edit the document, or to change or remove protection.|
| _UseIRM_|Optional|**Variant**|Specifies whether to use Information Rights Management (IRM) when protecting the document from changes.|
| _EnforceStyleLock_|Optional|**Variant**|Specifies whether formatting restrictions are enforced for a protected document.|
| _Type_|Required|WDPROTECTIONTYPE||
| _NoReset_|Optional|**Variant**||
| _Password_|Optional|**Variant**||
| _UseIRM_|Optional|**Variant**||
| _EnforceStyleLock_|Optional|**Variant**||

## Return value

**VOID**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

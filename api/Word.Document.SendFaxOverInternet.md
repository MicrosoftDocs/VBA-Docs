---
title: Document.SendFaxOverInternet method (Word)
keywords: vbawd10.chm158007760
f1_keywords:
- vbawd10.chm158007760
ms.prod: word
api_name:
- Word.Document.SendFaxOverInternet
ms.assetid: 1e1d061e-c33a-fdf1-ae63-b9a62babc1ef
ms.date: 06/08/2019
localization_priority: Normal
---

# Document.SendFaxOverInternet method (Word)

Sends a document to a fax service provider, who faxes the document to one or more specified recipients.

## Syntax

_expression_. `SendFaxOverInternet`( `_Recipients_` , `_Subject_` , `_ShowMessage_` )

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Recipients_|Optional| **Variant**|A **String** that represents the fax numbers and email addresses of the people to whom to send the fax. Separate multiple recipients with a semicolon.|
| _Subject_|Optional| **Variant**|A **String** that represents the subject line for the faxed document.|
| _ShowMessage_|Optional| **Variant**| **True** displays the fax message before sending it. **False** sends the fax without displaying the fax message.|

## Remarks

Using the **SendFaxOverInternet** method requires that a fax service is enabled on a user's computer. If a fax service is not enabled, the **SendFaxOverInternet** method will cause a run-time error.

The format used for specifying fax numbers in the Recipients parameter is either recipientsfaxnumber@usersfaxprovider or recipientsname@recipientsfaxnumber. You can access the user's fax provider information using the following registry path:

`HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Common\Services\Fax`

Use the FaxAddress key value at this registry location to determine the format to use for a user. If this registry entry does not exist, no fax service is available.

## Example

The following example sends a fax to the fax service provider, who will fax the message to the recipient.

```vb
ActiveDocument.SendFaxOverInternet _ 
 "14255550101@consolidatedmessenger.com", _ 
 "For your review", True
```

## See also

[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

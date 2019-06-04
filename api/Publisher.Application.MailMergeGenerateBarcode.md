---
title: Application.MailMergeGenerateBarcode event (Publisher)
keywords: vbapb10.chm268435489
f1_keywords:
- vbapb10.chm268435489
ms.prod: publisher
api_name:
- Publisher.Application.MailMergeGenerateBarcode
ms.assetid: 5da4ec65-32b6-ea05-09ad-d2224eafee30
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.MailMergeGenerateBarcode event (Publisher)

Occurs when Microsoft Publisher requires data to generate barcodes in a mail-merge publication, in particular when the mail-merge recipient list changes.


## Syntax

_expression_.**MailMergeGenerateBarcode** (_Doc_, _bstrString_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Doc_|Required| **Document**|The current publication.|
|_bstrString_|Required| **String**|Output parameter. A string representation of the barcode.|

## Remarks

Third-party add-ins that validate mail-merge addresses can use the **MailMergeGenerateBarcode** event to listen for user actions requesting that barcodes be generated. In this situation, when the add-in receives notification that the **MailMergeGenerateBarcode** event fired, and if the active document is connected to a data source, the add-in can use the **[MailMergeDataSource.ActiveRecord](Publisher.MailMergeDataSource.ActiveRecord.md)** property to determine the record for which to generate the barcode. If the active document is not connected to a data source, the add-in uses the address text directly.

If the add-in can use the address text directly, it returns a string representation of the barcode for the _bstrString_ output parameter. If the add-in cannot use the address text directly, it returns an empty string.

To permit triggering of the **MailMergeGenerateBarcode** event, you must handle the **[MailMergeInsertBarcode](Publisher.Application.MailMergeInsertBarcode.md)** event in your code, and the add-in must set the _OkToInsert_ parameter passed to that event to **True**. 

For more information about using events with the **Application** object, see [Using events with the Application object](../publisher/Concepts/using-events-with-the-application-object-publisher.md).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to handle the **MailMergeGenerateBarcode** event. It returns the string that represents the barcode for the active record. Note that the variable _indexNumberOfBarcodeColumn_ represents the index number of the column in the data source that lists barcodes. This code assumes that the current publication is connected to a data source.

```vb
Private Sub pubApplication_MailMergeGenerateBarcode(ByVal Doc As Document, bstrString As String) 
 bstrString = pubApplication.ActiveDocument.MailMerge.DataSource.DataFields.Item(indexNumberOfBarcodeColumn).Value 
End Sub
```

<br/>

For this event to occur, you must place the following line of code in the General Declarations section of your module.

```vb
Public WithEvents pubApplication As Application
```

<br/>

You then must run the following initialization procedure.

```vb
Public Sub Initialize_pubApplication() 
 Set pubApplication = Publisher.Application 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
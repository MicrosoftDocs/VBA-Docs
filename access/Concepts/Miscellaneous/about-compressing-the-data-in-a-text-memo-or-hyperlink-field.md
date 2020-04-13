---
title: About compressing the data in a Text, Memo, or Hyperlink field
ROBOTS: INDEX
keywords: vbaac10.chm1043024
f1_keywords:
- vbaac10.chm1043024
ms.prod: access
ms.assetid: 4f2aa9c5-a23e-076a-cc91-8b6061911f59
ms.date: 06/08/2019
localization_priority: Normal
---


# About compressing the data in a Text, Memo, or Hyperlink field

 
**Applies to:** Access 2013 | Access 2016

Microsoft Access uses the Unicode character-encoding scheme to represent the data in a Text, Memo, or Hyperlink field. Unicode represents each character as two bytes, so the data in a Text, Memo, or Hyperlink field requires more storage space than it did in Access 97 or earlier, where each character is represented as one byte.

To offset this effect of Unicode character representation and to ensure optimal performance, the default value of the **Unicode Compression** property for a Text, Memo, or Hyperlink field is **Yes**. When a field's **Unicode Compression** property is set to **Yes**, any character whose first byte is 0 is compressed when it is stored and uncompressed when it is retrieved. Because the first byte of a Latin character—a character of a Western European language such as English, Spanish, or German—is 0, Unicode character representation does not affect how much storage space is required for compressed data that consists entirely of Latin characters.

In a single field, you can store any combination of characters that Unicode supports. However, if the first byte of a particular character is not 0, that character is not compressed.

Data in a Memo field is not compressed unless it requires 4,096 bytes or less of storage space after compression. As a result, the contents of a Memo field might be compressed in one record, but might not be compressed in another record.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
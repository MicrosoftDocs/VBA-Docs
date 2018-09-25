---
title: Format Property - Text and Memo Data Types
keywords: vbaac10.chm5187266
f1_keywords:
- vbaac10.chm5187266
ms.prod: access
ms.assetid: 9d3c4e62-9328-28f2-da73-93c6277e11e3
ms.date: 06/08/2017
---


# Format Property - Text and Memo Data Types

**Applies to:** Access 2013 | Access 2016

You can use special symbols in the setting for the **Format** property to create custom formats for Text and Memo fields.


## Setting

You can create custom text and memo formats by using the following symbols.

|**Symbol**|**Description**|
|:-----|:-----|
|@|Text character (either a character or a space) is required.|
|&;|Text character is not required.|
|<|Force all characters to lowercase.|
|>|Force all characters to uppercase.|

Custom formats for Text and Memo fields can have up to two sections. Each section contains the format specification for different data in a field.

|**Section**|**Description**|
|:-----|:-----|
|First|Format for fields with text.|
|Second|Format for fields with zero-length strings and **Null** values.|

For example, if you have a text box control in which you want the word "None" to appear when there is no string in the field, you could type the custom format **@;"None"** as the control's **Format** property setting. The @ symbol causes the text from the field to be displayed; the second section causes the word "None" to appear when there is a zero-length string or Null value in the field.


> [!NOTE] 
> You can use the **Format** function to return one value for a zero-length string and another for a **Null** value, and you can similarly use the **Format** property to automatically format fields in table Datasheet view or controls on a form or report.


## Example

The following are examples of text and memo custom formats.

|**Setting**|**Data**|**Display**|
|:-----|:-----|:-----|
|@@@-@@-@@@@|465043799|465-04-3799|
|@@@@@@@@@|465-04-3799 465043799|465-04-3799 465043799|
|>|davolio DAVOLIO Davolio|DAVOLIO DAVOLIO DAVOLIO|
|<|davolio DAVOLIO Davolio|davolio davolio davolio|
|@;"Unknown"|**Null** value|Unknown|
||Zero-length string|Unknown|
||Any text| _Same text as entered is displayed_|


## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Search for specific Access error codes on Bing](https://www.bing.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access wiki on UtterAcess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)




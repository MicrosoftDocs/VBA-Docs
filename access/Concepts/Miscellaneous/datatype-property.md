---
title: DataType property
ROBOTS: INDEX
keywords: vbaac10.chm10085
f1_keywords:
- vbaac10.chm10085
ms.prod: access
api_name:
- Access.DataType
ms.assetid: 507dc426-afa4-783c-835d-5fdcb23a0e8d
ms.date: 06/08/2019
localization_priority: Priority
---


# DataType property

**Applies to:** Access 2013 | Access 2016

You can use the **DataType** property to specify the type of data stored in a table field. Each field can store data consisting of only a single data type.


## Setting

The **DataType** property uses the following settings.

|**Setting**|**Type of data**|**Size**|
|:-----|:-----|:-----|
|Text|**NOTE**: Text was replaced by Short text in Access 2013.(Default) Text or combinations of text and numbers, as well as numbers that don't require calculations, such as phone numbers. | Up to 255 characters or the length set by the **FieldSize** property, whichever is less. Microsoft Access does not reserve space for unused portions of a text field.|
|Short Text|**NOTE**: Short Text was introduced in Access 2013. It replaced Text.(Default) Text or combinations of text and numbers, as well as numbers that don't require calculations, such as phone numbers. | Up to 255 characters or the length set by the **FieldSize** property, whichever is less. Microsoft Access does not reserve space for unused portions of a text field.|
|Memo|**NOTE**: Memo was replaced by Long Text in Access 2013. Lengthy text or combinations of text and numbers. | Up to 63,999 characters. (If the Memo field is manipulated through DAO and only text and numbers [not binary data] will be stored in it, then the size of the Memo field is limited by the size of the database.)|
|Long Text|**NOTE**: Long Text was introduced in Access 2013. It replaced Memo. Lengthy text or combinations of text and numbers. | Up to 63,999 characters. (If the Long Text field is manipulated through DAO and only text and numbers [not binary data] will be stored in it, then the size of the Long Text field is limited by the size of the database.)|
|Number|Numeric data used in mathematical calculations. For more information on how to set the specific Number type, see the **FieldSize** property topic. | 1, 2, 4, or 8 bytes (16 bytes if the **FieldSize** property is set to Replication ID).|
|Date/Time|Date and time values for the years 100 through 9999. | 8 bytes.|
|Currency|Currency values and numeric data used in mathematical calculations involving data with one to four decimal places. Accurate to 15 digits on the left side of the decimal separator and to 4 digits on the right side. | 8 bytes.|
|AutoNumber|A unique sequential (incremented by 1) number or random number assigned by Microsoft Access whenever a new record is added to a table. AutoNumber fields can't be updated. For more information, see the **NewValues** property topic. | 4 bytes (16 bytes if the **FieldSize** property is set to Replication ID).|
|Yes/No|Yes and No values and fields that contain only one of two values (Yes/No, True **/** False, or On/Off). | 1 bit.|
|OLE Object|An object (such as a Microsoft Excel spreadsheet, a Microsoft Word document, graphics, sounds, or other binary data) linked to or embedded in a Microsoft Access table. | Up to 1 gigabyte (limited by available disk space)|
|Hyperlink|Text or combinations of text and numbers stored as text and used as a hyperlink address. A hyperlink address can have up to four parts: _text to display_ — the text that appears in a field or control; _address_ — the path to a file (UNC path) or page (URL); _subaddress_ — a location within the file or page; _screentip_ — the text displayed as a tooltip. | Each part of the parts of a Hyperlink data type can contain up to 2048 characters.|
|Attachment|Any supported type of file|You can attach images, spreadsheet files, documents, charts, and other types of supported files to the records in your database, much like you attach files to email messages. You can also view and edit attached files, depending on how the database designer sets up the Attachment field. Attachment fields provide greater flexibility than OLE Object fields, and they use storage space more efficiently because they don't create a bitmap image of the original file.|
|Lookup Wizard|Creates a field that allows you to choose a value from another table or from a list of values by using a list box or combo box. Clicking this option starts the Lookup Wizard, which creates a Lookup field. After you complete the wizard, Microsoft Access sets the data type based on the values selected in the wizard.|The same size as the primary key field used to perform the lookup, typically 4 bytes.|

You can set this property only in the upper portion of table Design view.

In Visual Basic , you can use the ADO **Type** property to set a field's data type before appending it to the **Fields** collection.


## Remarks

Memo, Hyperlink, and OLE Object fields can't be indexed.

> [!TIP] 
> Use the Currency data type for a field requiring many calculations involving data with one to four decimal places. **Single** and **Double** data type fields require floating-point calculation. The Currency data type uses a faster fixed-point calculation.

> [!WARNING] 
> Changing a field's data type after you enter data in a table causes a potentially lengthy process of data conversion when you save the table. If the data type in a field conflicts with a changed **DataType** property setting, you may lose some data.

Set the **Format** property to specify a predefined display format for Number, Date/Time, Currency, and Yes/No data types.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

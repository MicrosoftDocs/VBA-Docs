---
title: Format Property - Yes/No Data Type
keywords: vbaac10.chm5187268
f1_keywords:
- vbaac10.chm5187268
ms.prod: access
ms.assetid: 51b9af9b-8c43-8f3a-cf93-fc0f3a7eb0a5
ms.date: 06/08/2017
---


# Format Property - Yes/No Data Type

**Applies to:** Access 2013 | Access 2016

You can set the **Format** property to the Yes/No, **True** / **False**, or On/Off predefined formats or to a custom format for the Yes/No data type.

## Setting
<a name="sectionSection0"> </a>

Microsoft Access uses a check box control as the default control for the Yes/No data type. Predefined and custom formats are ignored when a check box control is used. Therefore, these formats apply only to data that is displayed in a text box control.


## Predefined Formats
<a name="sectionSection1"> </a>

Yes, **True**, and On are equivalent, as are No, **False**, and Off. If you specify one predefined format and then enter an equivalent value, the predefined format of equivalent value will be displayed. For example, if you enter **True** or On in a text box control with its **Format** property set to Yes/No, the value is automatically converted to Yes.


## Custom Formats
<a name="sectionSection2"> </a>

The Yes/No data type can use custom formats containing up to three sections.

|**Section**|**Description**|
|:-----|:-----|
|First|This section has no effect on the Yes/No data type. However, a semicolon (;) is required as a placeholder.|
|Second|The text to display in place of Yes, **True**, or On values.|
|Third|The text to display in place of No, **False**, or Off values.|

## Example
<a name="sectionSection3"> </a>

The following example shows a custom yes/no format for a text box control. The control displays the word "Always" in blue text for Yes, **True**, or On, and the word "Never" in red text for No, **False**, or Off.

```vb
;"Always"[Blue];"Never"[Red]
```

## See also

- [Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Search for specific Access error codes on Bing](https://www.bing.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access wiki on UtterAcess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)



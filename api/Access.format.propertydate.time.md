---
title: Date/Time data type (Format property)
ms.prod: access
ms.assetid: d043c816-aefe-4881-90bd-59dcbb3b28da
ms.date: 11/29/2018
localization_priority: Normal
---


# Date/Time data type (Format property)

**Applies to:** Access 2013 | Access 2016

You can set the **Format** property to predefined date and time formats or use custom formats for the Date/Time data type.

## Settings

### Predefined formats

The following table shows the predefined **Format** property settings for the Date/Time data type.

|Setting|Description|
|:------|:----------|
|General Date|(Default) If the value is a date only, no time is displayed; if the value is a time only, no date is displayed. This setting is a combination of the Short Date and Long Time settings.<br/><br/>Examples: 4/3/93, 05:34:00 PM, and 4/3/93 05:34:00 PM.|
|Long Date|Same as the Long Date setting in the regional settings of Windows.<br/><br/>Example: Saturday, April 3, 1993.|
|Medium Date|Example: 3-Apr-93.|
|Short Date|Same as the Short Date setting in the regional settings of Windows.<br/><br/>Example: 4/3/93.<br/><br/>**CAUTION**: The Short Date setting assumes that dates between 1/1/00 and 12/31/29 are twenty-first century dates (that is, the years are assumed to be 2000 to 2029). Dates between 1/1/30 and 12/31/99 are assumed to be twentieth century dates (that is, the years are assumed to be 1930 to 1999).|
|Long Time|Same as the setting on the **Time** tab in the regional settings of Windows.<br/><br/>Example: 5:34:23 PM.|
|Medium Time|Example: 5:34 PM.|
|Short Time|Example: 17:34.|

### Custom formats

You can create custom date and time formats by using the following symbols.

|Symbol|Description|
|:-----|:----------|
|: (colon)|Time separator. Separators are set in the regional settings of Windows.|
|/|Date separator.|
|c|Same as the General Date predefined format.|
|d|Day of the month in one or two numeric digits, as needed (1 to 31).|
|dd|Day of the month in two numeric digits (01 to 31).|
|ddd|First three letters of the weekday (Sun to Sat).|
|dddd|Full name of the weekday (Sunday to Saturday).|
|ddddd|Same as the Short Date predefined format.|
|dddddd|Same as the Long Date predefined format.|
|w|Day of the week (1 to 7).|
|ww|Week of the year (1 to 53).|
|m|Month of the year in one or two numeric digits, as needed (1 to 12).|
|mm|Month of the year in two numeric digits (01 to 12).|
|mmm|First three letters of the month (Jan to Dec).|
|mmmm|Full name of the month (January to December).|
|q|Date displayed as the quarter of the year (1 to 4).|
|y|Number of the day of the year (1 to 366).|
|yy|Last two digits of the year (01 to 99).|
|yyyy|Full year (0100 to 9999).|
|h|Hour in one or two digits, as needed (0 to 23).|
|hh|Hour in two digits (00 to 23).|
|n|Minute in one or two digits, as needed (0 to 59).|
|nn|Minute in two digits (00 to 59).|
|s|Second in one or two digits, as needed (0 to 59).|
|ss|Second in two digits (00 to 59).|
|ttttt|Same as the Long Time predefined format.|
|AM/PM|Twelve-hour clock with the uppercase letters "AM" or "PM" as appropriate.|
|am/pm|Twelve-hour clock with the lowercase letters "am" or "pm" as appropriate.|
|A/P|Twelve-hour clock with the uppercase letter "A" or "P" as appropriate.|
|a/p|Twelve-hour clock with the lowercase letter "a" or "p" as appropriate.|
|AMPM|Twelve-hour clock with the appropriate morning/afternoon designator as defined in the regional settings of Windows.|

Custom formats are displayed according to the settings specified in the regional settings of Windows. Custom formats inconsistent with the settings specified in the regional settings of Windows are ignored.

> [!NOTE] 
> If you want to add a comma or other separator to a custom format, enclose the separator in quotation marks as follows: `mmm d", "yyyy`.


## Example

The following are examples of custom date/time formats.

|Setting|Display|
|:-----|:-----|
|`ddd", "mmm d", "yyyy`|Mon, Jun 2, 1997|
|`mmmm dd", "yyyy`|June 02, 1997|
|`"This is week number "ww`|This is week number 22|
|`"Today is "dddd`|Today is Tuesday|

You could use a custom format to display "A.D." before or "B.C." after a year depending on whether a positive or negative number is entered. To see this custom format work, create a new table field, set its data type to Number, and enter a format as follows:

`"A.D. " #;# " B.C."`

Positive numbers are displayed as years with an "A.D." before the year. Negative numbers are displayed as years with a "B.C." after the year.


## See also

- [Number and Currency](Access.format.propertynumber.and.currency.md)
- [Text and Memo](Access.format.propertytext.and.memo.md)
- [Yes/No](Access.format.propertyyes.no.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

---
title: Number and Currency data types (Format property)
ms.prod: access
ms.assetid: f48fbfad-c249-4011-9b3e-bbd6628ac1f7
ms.date: 11/29/2018
localization_priority: Priority
---


# Number and Currency data types (Format property)

**Applies to:** Access 2013 | Access 2016

You can set the **Format** property to predefined number formats or custom number formats for the Number and Currency data types.

## Settings

### Predefined formats

The following table shows the predefined **Format** property settings for numbers.

|Setting|Description|
|:-----|:-----|
|General Number|(Default) Display the number as entered.|
|Currency|Use the thousand separator; follow the settings specified in the regional settings of Windows for negative amounts, decimal and currency symbols, and decimal places.|
|Euro|Use the euro symbol (![Euro symbol](../images/euro_ZA06048440.gif)), regardless of the currency symbol specified in the regional settings of Windows.|
|Fixed|Display at least one digit; follow the settings specified in the regional settings of Windows for negative amounts, decimal and currency symbols, and decimal places.|
|Standard|Use the thousand separator; follow the settings specified in the regional settings of Windows for negative amounts, decimal symbols, and decimal places.|
|Percent|Multiply the value by 100 and append a percent sign (%); follow the settings specified in the regional settings of Windows for negative amounts, decimal symbols, and decimal places.|
|Scientific|Use standard scientific notation.|

### Custom formats

Custom number formats can have one to four sections with semicolons (;) as the list separator. Each section contains the format specification for a different type of number.

|Section|Description|
|:-----|:-----|
|First|The format for positive numbers.|
|Second|The format for negative numbers.|
|Third|The format for zero values.|
|Fourth|The format for **Null** values.|

For example, you could use the following custom Currency format:

```vb
$#,##0.00[Green];($#,##0.00) [Red];"Zero";"Null"
```

This number format contains four sections separated by semicolons and uses a different format for each section.

If you use multiple sections but don't specify a format for each section, entries for which there is no format will either display nothing or default to the formatting of the first section.

You can create custom number formats by using the following symbols.

|Symbol|Description|
|:-----|:-----|
|`.` (period)|Decimal separator. Separators are set in the regional settings in Windows.|
|`,` (comma)|Thousand separator.|
|`0`|Digit placeholder. Display a digit or 0.|
|`#`|Digit placeholder. Display a digit or nothing.|
|`$`|Display the literal character "$".|
|`%`|Percentage. The value is multiplied by 100 and a percent sign is appended.|
|`E-` or `e-`|Scientific notation with a minus sign (-) next to negative exponents and nothing next to positive exponents. This symbol must be used with other symbols, as in 0.00E-00 or 0.00E00.|
|`E+` or `e+`|Scientific notation with a minus sign (-) next to negative exponents and a plus sign (+) next to positive exponents. This symbol must be used with other symbols, as in 0.00E+00.|

## Remarks

You can use the **DecimalPlaces** property to override the default number of decimal places for the predefined format specified for the **Format** property.

The predefined currency and euro formats follow the settings in the regional settings of Windows. You can override these by entering your own currency format.


## Examples

The following are examples of the predefined number formats.

|Setting|Data|Display|
|:-----|:-----|:-----|
|General Number|3456.789 -3456.789 $213.21|3456.789 -3456.789 $213.21|
|Currency|3456.789 -3456.789|$3,456.79 ($3,456.79)|
|Fixed|3456.789 -3456.789 3.56645|3456.79 -3456.79 3.57|
|Standard|3456.789|3,456.79|
|Percent|3 0.45|300% 45%|
|Scientific|3456.789 -3456.789|3.46E+03 -3.46E+03|

The following are examples of custom number formats.

|Setting|Description|
|:-----|:-----|
|`0;(0);;"Null"`|Display positive values normally; display negative values in parentheses; display the word "Null" if the value is **Null**.|
|`+0.0;-0.0;0.0`|Display a plus (+) or minus (-) sign with positive or negative numbers; display 0.0 if the value is zero.|

## See also

- [Date/Time](Access.format.propertydate.time.md)
- [Text and Memo](Access.format.propertytext.and.memo.md)
- [Yes/No](Access.format.propertyyes.no.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
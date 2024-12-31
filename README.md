# Lightweight - XLSX Util
A Lightweight Salesforce Apex utility to read and write basic XLSX files based on the ECMA-376 standard fully on platform without the need for any external API calls for processing. The application is based on the `ZipWriter` and `ZipReader` classes from the compression namespace. These are currently in developer preview but should go GA in the Sprint 25 release.

We are limited to the Apex govenor limits, but nonetheless it can write pretty large worksheets synchronously and even larger asynchronously. If you need extremly large or complex files, you're probably better off using Salesforce Document Builder or a 3rd party application like Conga Composer.

The library comes with the most common functionalities like freezing rows, merging cells, create hyperlinks and the majority of the styling options. There is no support for table styles, charts or images at this time.

The application structure is built in a flexible way so that if you need to extend the application with specific functions, you can easily add additional XML elements as per the ECMA-376 - Office Open XML standard.

Because of the potential very large number of cells a few trade-off have been made in consistent method usage. But they are roughly the same.

The XML file structures are not rocket science, but for a file to be valid and not give you any errors it can be very (case) sensitive. If you make changes make sure to follow the standards exactly. Friendly warning: A single wrong capital in an XML element name can break your file, so test regularly.

## Blog
- Salesforce.com Developer Blog:
- 

## Package Info
| Info | Value |
|---|---|
|Name|Lightweight - XLSX Builder|
|Version|0.1.0-1|
|Managed Installation URL | */packaging/installPackage.apexp?p0=xxx* |
|Unlocked Installation URL| */packaging/installPackage.apexp?p0=xxx* |

## Parse Excel files
Parsing is done using the `Parse` class in the `xlsx` namespace. We can parse to two different formats: a *multi dimensional array* or a *list of maps*. In the array format the first list represents the worksheet, the child the rows and the grand child the cells.
In the Map List option, each map represents a worksheet and the Map Key equals the cell name like "A1" and the value the corresponding value object.

The basic method outline is as below:
```java
// OUTPUT AS MULTIDIMENSIONAL ARRAY
Object[][][] xlsxDataArray = xlsx.Parse.toArray(Map<String,Compression.ZipEntry> entries){}

// OUTPUT AS MAP
List<Map<String,Object>> xlsxDataMap = xlsx.Parse.toMap(Map<String,Compression.ZipEntry> entries){}
```

The entries parameter is the unzipped blob data represented in a map with entries `Map<String,Compression.ZipEntry>`. The blob data can come from a file, attachment or web service. As long as it is a valid XLSX file.
```java
// The document Id
Id documentId = '015Qz000004jf7yIAA';

// Query a document for content, this is a compressed zip file body
Blob xlsxBlobData = [SELECT body FROM Document WHERE Id = :documentId LIMIT 1]?.Body;

// Create a new zip reader instance
Compression.ZipReader reader = new Compression.ZipReader(xlsxBlobData);

// Get a map with entries
Map<String,Compression.ZipEntry> entries = reader.getEntriesMap();
```

To preserve heap size this can be simplyfied by putting all statements inline instead of separate variables:
```java
// As multi dimensional array
Object[][][] xlsxDataArray = xlsx.Parse.toArray(
    new Compression.ZipReader(
        [SELECT body FROM Document WHERE Id = '015Qz000004jf7yIAA' LIMIT 1]?.Body
    ).getEntriesMap()
);


// As data map
List<Map<String,Object>> xlsxDataMap = xlsx.Parse.toMap(
    new Compression.ZipReader(
        [SELECT body FROM Document WHERE Id = '015Qz000004jf7yIAA' LIMIT 1]?.Body
    ).getEntriesMap()
);
```

## Parse Methods
For different use cases you can use different parse methods each with advantages. Using the `Dom.Document` class for reading XML is a lot faster than the `XmlStreamWriter` but also limited due to the large heap size it uses.
The `xlsx.Parse` class is used to parse an XLSX file body from an unzipped file body

|Return type| Method signature| Use for |
|---|---|---|
| `Object[][][]`             |`xlsx.Parse.toArray(Map<String,Compression.ZipEntry> entries)`                                    | Large files, CSV like data |
| `Object[][][]`             |`xlsx.Parse.toArray(Map<String,Compression.ZipEntry> entries, Set<Integer> selectedSheets)`       | Large files, CSV like data, specific sheets only |
| `Object[][][]`             |`xlsx.Parse.toArrayDomDoc(Map<String,Compression.ZipEntry> entries)`                              | Small files, CSV like data  |
| `Object[][][]`             |`xlsx.Parse.toArrayDomDoc(Map<String,Compression.ZipEntry> entries, Set<Integer> selectedSheets)` | Small files, CSV like data, specific sheets only |
| `List<Map<String,Object>>` |`xlsx.Parse.toMap(Map<String,Compression.ZipEntry> entries)`                                      | Large files, Data based on cell names |
| `List<Map<String,Object>>` |`xlsx.Parse.toMap(Map<String,Compression.ZipEntry> entries, Set<Integer> selectedSheets)`         | Large files, Data based on cell names, specific sheets only |
| `List<Map<String,Object>>` |`xlsx.Parse.toMapDomDoc(Map<String,Compression.ZipEntry> entries)`                                | Small files, Data based Cell Name |
| `List<Map<String,Object>>` |`xlsx.Parse.toMapDomDoc(Map<String,Compression.ZipEntry> entries, Set<Integer> selectedSheets)`   | Small files, Data based on cell namese, specific sheets only |
| `Map<String,Integer>`      |`xlsx.Parse.toWorksheetNameIndexMap(Map<String,Compression.ZipEntry> entries)`                    | You need to get the index based on the name of the worksheet |


## Build Methods
The `xlsx.Build` class is used to create an XLSX file from a `xlsx.Builder` class instance.
|Return type| Method signature| Use for |
|---|---|---|
|`Document`         |`xlsx.Build.asDocument(Builder b)`      | Building your XLSX file as a `Document` Object|
|`ContentVersion`   |`xlsx.Build.asContentVersion(Builder b)`| Building your XLSX file as a `ContentVersion` Object|


## Builder Methods
The `xlsx.Builder` class is used to setup/configure the entire XLSX document. Use this to add worksheets, rows and cells but also configure the document name etc.
Making changes to anything in the sheets uses the worksheet index (`wi`), column index (`ci`) and row index (`ri`) parameters. These are the zero indexes so cell A1 on Sheet1 is set as [0][0][0].
|Return type| Method signature| Use for |
|---|---|---|
|`void`     |`setUseSharedStrings(Boolean value)`                    | Enable or disable shared strings. Shared strings is recommended and enabled by default.|
|`void`     |`setIncludeDefaultStyles(Boolean value)`                | Enable the Lightweight - XLSX Util standard styles |
|`void`     |`setFileName(String value)`                             | Setting a custom name for the output file, does not have to end in .xlsx|
|`void`     |`setTitle(String value)`                                | Setting a document title|
|`void`     |`setSubject(String value)`                              | Setting a document subject|
|`void`     |`setDescription(String value)`                          | Setting a document description|
|`void`     |`addKeyword(String value)`                              | Adding a document keyword, A max of 255 characters for all keywords is in place|
|`Integer`  |`addWorksheet(String name)`                             | Adds a new worksheet with a name (max 31 characters). Set the name param to NULL for an auto generated name like in Excel (i.e. sheet1, sheet2 etc.) |
|`void`     |`prePadWorksheet(Integer wi, Integer ci, Integer ri)`   | If you know the max number of rows and columns you can use this method for performance optimization|
|`void`     |`setVisible(Integer wi, Boolean state)`                        | Set a worksheet visible or invisible|
|`void`     |`setTabColor(Integer wi, String colorCode)`                    | Set a custom tab color (6 digit HEX code)|
|`void`     |`setAutoFilter(Integer wi, Boolean value)`                     | Enable an auto filter on the columns|
|`void`     |`setFreezeRows(Integer wi, Integer numberOfRows)`              | Set a number of rows to freeze on a worksheet|
|`void`     |`setFreezeColumns(Integer wi, Integer numberOfColumns)`        | Set a number of columns to freeze on a worksheet|
|`void`     |`addTextCell(Integer wi, Integer ci, Integer ri, String v)`                | Add a Text cell|
|`void`     |`addNumberCell(Integer wi, Integer ci, Integer ri, Integer v)`             | Add an Integer cell|
|`void`     |`addNumberCell(Integer wi, Integer ci, Integer ri, Decimal v)`             | Add a Decimal cell|
|`void`     |`addBooleanCell(Integer wi, Integer ci, Integer ri, Boolean v)`            | Add a Boolean cell|
|`void`     |`addFormulaCell(Integer wi, Integer ci, Integer ri, Object v, String f)`   | Add a Formula cell (text for value)|
|`void`     |`addTextCell(Integer wi, Integer ci, Integer ri, String v, Integer s)`             | Add a Text cell with style index|
|`void`     |`addNumberCell(Integer wi, Integer ci, Integer ri, Integer v, Integer s)`          | Add an Integer cell with style index|
|`void`     |`addNumberCell(Integer wi, Integer ci, Integer ri, Decimal v, Integer s)`          | Add a Decimal cell with style index|
|`void`     |`addBooleanCell(Integer wi, Integer ci, Integer ri, Boolean v, Integer s)`         | Add a Boolean cell with style index|
|`void`     |`addFormulaCell(Integer wi, Integer ci, Integer ri, Object v, String f, Integer s)`| Add a Formula cell with style index|
|`void`     |`addMergeCell(Integer wi, Integer startCi, Integer startRi, Integer endCi, Integer endRi))`| Merge cells together. Note: merge cells cannot overlap or your sheet won't work.|
|`void`     |`addHyperLink(Integer wi, Integer ci, Integer ri, String location, String display)`        | Add a hyperlinked cell, note you need to add a text cell first before adding a hyperlink|
|`void`     |`setRowStyle(Integer wi, Integer ri, Integer s)`               | Set a row's style index |
|`void`     |`void setRowHeight(Integer wi, Integer ri, Decimal h)`         | Set a row's height |
|`void`     |`setRowHidden(Integer wi, Integer ri, Boolean v)`              | Set a row's visibility |
|`void`     |`setColStyle(Integer wi, Integer ci, Integer s)`               | Set a column's style index |
|`void`     |`setColWidth(Integer wi, Integer ci, Decimal w)`               | Set a column's height |
|`void`     |`setColHidden(Integer wi, Integer ci, Boolean h)`              | Set a column's visibility|
|`void`     |`setCellStyle(Integer wi, Integer ci, Integer ri,Integer s)`   | Set a cell's style after it has been created |

## Styles Methods
The `xlsx.StylesBuilder` class is used to create custom styles and add them to a `xlsx.Builder` class instance.
The `add methods` return an `index`, the indexes for number formats, fonts, files borders and alignments need to be used as the input parameters for the `addCellStyle()` method to create a unique style. Indexes can be reused to mix and match styles.
|Return type| Method signature| Use for |
|---|---|---|
|`Integer`              |`addNumberFormat(Builder b, Integer numFmtId, String formatCode)`     |Add a custom number format, you'll rarely need this one|
|`Integer`              |`addFont(Builder b, Integer sz, String name, String rgb, Boolean bold, Boolean italic, Boolean underline)` | Add a custom font with size color and decoration|
|`Integer`              |`addFill(Builder b, String patternType, String fgColor, String bgColor)`   |Add a fill with a pattern, foreground and background color|
|`Map<String,String>`   |`borderConfig(String style, String color)`     |Attribute for the `addBorder` method|
|`Integer`              |`addBorder(Builder b, Map<String,String> left, Map<String,String> right, Map<String,String> top, Map<String,String> bottom)`   | Add a custom border with a style and a color. Valid values are: |
|`Integer`              |`addAlignment(Builder b, String horizontal, String vertical, Integer textRotation, Boolean wrapText)`  | Add a custom alignment for the cell. Valid values are: |
|`Integer`              |`addCellStyle(Builder b, Integer numFmtId, Integer fontId, Integer fillId, Integer borderId, Integer alignmentId)` | Combine the indexes from previous methods to create a unique style index that can be used in for row, columns and cells.|
|`Integer`              |`getHeaderStyleIndex(Integer ci, Integer startCi, Integer endCi)`  | If you include the standard styles, use this method to get the index for a header of a "table". Set the ci, the start ci of teh table and the last ci of the table.|
|`Integer`              |`getMainStyleIndex(Integer ci, Integer ri, Integer startCi, Integer endCi, Integer endRi)` |If you include the standard styles, use this method to get the index for the rows (including the bottom row). Set the first and last indexes of the cell range. |


## CommonUtil Methods
The `xlsx.CommonUtil` class contains utilities that help you with the setup of a `xlsx.Builder` class instance. 
|Return type| Method signature| Use for |
|---|---|---|
|`String`     | `columnNameFromColumnIndex(Integer columnIndex)`   |Translate a column `index` to a column name. I.e. 0=A and 4=E etc.|
|`Integer`    | `columnNumberFromColumnName(String columnName)`    |Get a column `number` from a column name. I.e A=1 and E=5|
|`String`     | `cellName(Integer columnIndex, Integer rowIndex)`  |Get cell name based on row and column `index`. Ie,e 0,1 = A1, 4,4 = E5 etc.|
|`String`     | `randomHtmlHexColorCode()`                         |Creates a random 6 digit HEX code that can be used in tabs. Handy for testing purposes to distinguish between tabs.|
|`Datetime`   | `getTimestamp()`                                   |Get the timestamp that is used when creating an `xlsx.Builder` class instance. This ensures you use the same timestamp through out.|
|`String`     | `getTimestampString()`                             |Get the timestamp as a string that can be used in a file name|


## Exceptions
In order to handle Exceptions there are two Exception classes: `xlsx.ParseException` and `xlsx.BuildException`. The first are thrown on any parsing issues and the second one on anything related to building the XLSX file.


## Examples
The `examples folder` contains a number of examples you can check out:
### Parse Examples
-
-
-
-

### Build Examples
-
-
-
-

# Additional resources
- https://www.brandwares.com/downloads/Open-XML-Explained.pdf


# Getting started guide
## Set up the document properties
Each new file you create starts with an `xlsx.Builder` class instance. The builder is used to "set up" the entire file.
Let's start with a verbose example where we set the string handling, if we want to include some app specific default styles and the file properties.
```java
// Create a new builder
xlsx.Builder b = new xlsx.Builder();

// Because we have repetition for each sObject (field names),
// use shared strings instead of inline strings for a better overall performance
// Defaults to true, so not required
b.setUseSharedStrings(true);

// This option createss the default set of "table" styles
// Can be used to use a default styling for all sheets, (will take a small performance hit)
// Defaults to false
b.setIncludeDefaultStyles(true);
 
// Set the name of the xlsx file
b.setFileName(xlsx.CommonUtil.getTimestampString() + '_example.xlsx' );

// Set file properties
b.setTitle('Example ('+xlsx.CommonUtil.getTimestampString()+ ')');
b.setSubject('Example subject');
b.setDescription('Example description');
b.addKeyword('Example');
b.addKeyword('Additional Example');
```

## Setup worksheets
Data is managed by using a multi-dimentional array with zero based index for all worksheets. Each worksheet has a worksheet index (`wi`), column index (`ci`) and row index (`ri`).
This means that for example in Apex a cell located at [0][0][0] is located at "A1" in the first worksheet and [1][1][1] is located at "B2" in the second worksheet etc. This makes coding and array looping a whole lot easier. But it is something to keep in mind when adding data to the grid.
I tried to match it to Excel number at first, but the +1/-1 drove me completely insane, so I reverted back to "good old starting at 0".

So now that we have our file set up, lets start adding worksheets. We add a worksheet using the `addWorksheet(String name)` method. This method return an integer with the sheet index (`wi`) of the sheet you have added. In case you don't want to manually keep track you can add the index to a variable.
Worksheet names cannot have funky characters and not be longer than 31 characters. I have put some sanitization logic in place, but that might not be completely fool proof.

```java
// Get the worksheet index when adding it to the worksheet array
Integer wi0 = b.addWorksheet('My First Worksheet');
Integer ws1 = b.addWorksheet('My Second Worksheet');

// Give the worksheet tabs a nice color (These are are the Data Cloud Logo Colors in case you wonder)
// You use a HEX code and the tab index
b.setTabColor(wi0,'5E62A2');
b.setTabColor(wi1,'868ACA');

// We can freeze rows and columns by specifying the wi and the number of rows or
// columns we want to freeze
// Freeze top row
b.freezeRows(wi0, 1);
b.freezeRows(wsi1, 1);

// Freeze the left column
b.freezeColumns(wi0, 1);
b.freezeColumns(wi1, 1);

// Enable the filter options for the first row on both sheets
// You'll see that the filter will be applied for each column added
// It cannot be more specific, it's all or none.
b.enableAutoFilter(wi0);
b.enableAutoFilter(wi1);
```

## Add Cells

## Add Merge Cells and Hyperlinks

## Add Styling

## Build as Document or ContentVersion




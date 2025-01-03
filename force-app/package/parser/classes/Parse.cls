/**
 * @author         Justus van den Berg (jfwberg@gmail.com)
 * @date           January 2025
 * @copyright      (c) 2025 Justus van den Berg
 * @license        MIT (See LICENSE file in the project root)
 * @description    A class to holds an example of how you can parse an XLSX
 *                 worksheet data into a multi-dimensional array or into a list of maps
 *
 * @false-positive AvoidGlobalModifier      This is a utility that is designed to be called from a
 *                                          managed package. It can be used for a user's own custom
 *                                          Implementation, so global is the way to open up this
 *                                          utility for global use.
 *
 * @parser rules   - The worksheet should have the "dimensions" element populated
 *                 - Formula fields are NOT recalculated, so they need to contain the correct value
 *                 - You work with raw data so hidden and (password) protected rows and columns will be parsed
 *                 - Password encrypted sheets are not supported
 *
 * @Notes          - There are no checks if the XLSX files are valid
 *                 - The Shared String Indexes are not tested for out of bounds violations due to performance
 */
@SuppressWarnings('PMD.AvoidGlobalModifier')
global with sharing class Parse{

    /** **************************************************************************************************** **
     **                                       GLOBAL WORKSHEET METHODS                                       **
     ** **************************************************************************************************** **/
    /**
     * @description Method that creates a map with the sheet name as key and the
     *              sheet index as the value
     *              This is to be used if you dont know the sheet order but do know the name
     * @param  entries A map of path/zip entries from the unzipped XLSX file
     * @return         A map with the sheet name as key and the index as value
     */
    global static Map<String,Integer> toWorksheetNameIndexMap(Map<String,Compression.ZipEntry> entries){

        // Validate the workbook file exists
        if(!entries.containsKey(Constant.ZP_PATH_WORKBOOK)){
            throw new ParseException(String.format(Constant.MSG_INVALID_ZIP_ENTRY, new String[]{Constant.ZP_PATH_WORKBOOK}));
        }

        // List of sheet indexes to create to proper order
        SheetIndex[] sheetIndexes = new SheetIndex[]{};

        // Create an XSR for the current worksheet XML file
        XmlStreamReader xsr = new XmlStreamReader(entries.get(Constant.ZP_PATH_WORKBOOK).getContent().toString());

        // I am not using while xsr.hasNext() to save a method call at the end of each item
        while(true){

            // Start at the beginning of the book and make sure that it is a book
            if (xsr.getEventType() == XmlTag.START_ELEMENT) {

                // Get the sheet element and create a new SheetIndex object for each sheet so we can sort it
                switch on xsr.getLocalName(){
                    when 'sheet'{
                        sheetIndexes.add(
                            new SheetIndex(
                                Integer.valueOf(xsr.getAttributeValue(null, 'sheetId')),
                                xsr.getAttributeValue(null, 'name')
                            )
                        );
                    }
                }
            }

            // If there is a next element go to the next element, else we're done
            if(xsr.hasNext()){ xsr.next(); } else { break; }
        }

        // Sort the sheet indexes
        sheetIndexes.sort();

        // Create a map for the output
        Map<String,Integer> worksheetNameIndexMap = new Map<String,Integer>();

        // populate the output map
        for(Integer i =0,max=sheetIndexes.size();i<max;i++){
            worksheetNameIndexMap.put(sheetIndexes[i].name,i);
        }

        // Return the output map
        return worksheetNameIndexMap;
    }


    /** **************************************************************************************************** **
     **                                         GLOBAL ARRAY METHODS                                         **
     ** **************************************************************************************************** **/
    /**
     * @description Method that takes an unzipped xlsx archive as parameter and
     *              returns the sheet data in a multi dimentional array in the form of
     *              data[WorksheetIndex][ColumnIndex][RowIndex]
     * @param  entries A map of path/zip entries from the unzipped XLSX file
     * @return      A list of sheets with a list of rows with a list of cell
     * @note        Use this method for the large sheets
     */
    global static Object[][][] toArray(Map<String,Compression.ZipEntry> entries){
        return toArray(entries, null);
    }


    /**
     * @description Method that takes an unzipped xlsx archive as parameter and
     *              returns the sheet data in a multi dimentional array in the form of
     *              data[WorksheetIndex][ColumnIndex][RowIndex]
     * @param  entries        A map of path/zip entries from the unzipped XLSX file
     * @param  selectedSheets A set of integers representing the sheet indexes
     * @return      A list of sheets with a list of rows with a list of cells
     * @note        Use this method for the large sheets
     */
    global static Object[][][] toArray(Map<String,Compression.ZipEntry> entries, Set<Integer> selectedSheets){

        // Create a multi-dimentional array for the output
        // Worksheet - Column - Row
        Object[][][] output = new Object[][][]{};

        // Create a list of shared strings
        String[] sharedStrings = getSharedStrings(entries);

        // In order to sort the worksheets, we need to to put them in a list to sort
        String[] worksheetPaths = getWorksheetPaths(entries);

        // Keep track of the largest number of columns. If one sheet is larger than the other
        // we do not do have a duplicate translation for overlapping column names
        // i.e. If sheet 1 has 10 colums and sheet 2 has 20 columns, make sure to only
        // translate the column names from 11 to 20 to save some CPU time
        Integer maxColDimension = 0;

        // Keep track of skipped worksheets
        Integer worksheetSkipCount = 0;

        // Parse each worksheet in it's own array
        for(Integer wi =0, wiMax=worksheetPaths.size(); wi<wiMax; wi++){

            // Skip sheets that are not included in the selectedSheets
            if(selectedSheets != null && !selectedSheets.contains(wi)){
                worksheetSkipCount++;
                continue;
            }

            // Calculate the current worksheet index in case we select specific worksheets only
            Integer currentWi = wi-worksheetSkipCount;

            // Add array for the sheet to hold the columns and rows
            output.add(new Object[][]{});

            // Create an XSR for the current worksheet XML file
            XmlStreamReader xsr = new XmlStreamReader(entries.get(worksheetPaths[wi]).getContent().toString());

            // Set the current row to -1 so the indexes match
            Integer currentRi = -1;

            // Variables to hold the type and name of the current open cell element
            String cellType;
            String columnName;

            // Traverse the XML elements in the worksheet
            // I am not using while xsr.hasNext() to save a method call at the end of each item
            while(true) {

                // Start at the beginning of the book and make sure that it is a book
                if (xsr.getEventType() == XmlTag.START_ELEMENT) {
                    switch on xsr.getLocalName() {

                        // First we need the dimention of the worksheet to find out the max size of our rows
                        // The order in the XML is mandatory, so this value should exist, if not the sheet is invalid
                        when 'dimension'{

                            // Get the cell name from the range, this tag is mandatory so should be there
                            String cName   = xsr.getAttributeValueAt(0).split(':')[1];

                            // Get the ref element from the token and calculate the value
                            String colName = ParseUtil.columnNameFromCellName(cName);

                            // Get the number of columsn and rows that are in this worksheet
                            Integer numCols = CommonUtil.columnNumberFromColumnName(colName);
                            Integer numRows = Integer.valueOf(cName.removeStart(colName));

                            // Pre-populate a map with the column name related to the column index i.e. AA => 27, AB =>28 etc
                            // Prepopulating makes the parsing go a lot faster as we only have to calculate it once
                            if(numCols > maxColDimension){

                                // Popualte the colum index (startnumber and the total number of columns)
                                ParseUtil.populateColumnIndexMap(maxColDimension, numCols);

                                // Update to the new max column dimension
                                maxColDimension = numCols;
                            }

                            // Add empty values for each (potential cell)
                            for(Integer ci=0; ci<numCols; ci++){

                                // Pre-populate the entire "grid" with empty lists and null values
                                output[currentWi].add(new Object[]{});

                                // Create a new entry for the row
                                for(Integer ri=0; ri<numRows; ri++){
                                    output[currentWi][ci].add(null);
                                }
                            }
                        }

                        // The row element, keep track of what row we're currently in
                        // We can have a situation where we have row 1 and 10 populated, but nothing in between
                        when 'row'{
                            currentRi = Integer.valueOf(xsr.getAttributeValue(null, 'r'))-1;
                        }

                        // The cell element
                        when 'c' {
                            columnName = ParseUtil.columnNameFromCellName(xsr.getAttributeValue(null, 'r'));
                            cellType   = xsr.getAttributeValue(null, 't');
                        }

                        // Value tags inside the cell element
                        when 'v','t'{

                            // Skip to the text element
                            xsr.next();

                            // Validate we're in a text element
                            if(xsr.getEventType() == XmlTag.CHARACTERS){

                                // Put the cell value in array at location wi.ci.ri
                                // If there is a "sharedStrings index", get the shared strings value at that index
                                // else the value in the value in the XSR
                                output[currentWi][ParseUtil.columnIndexMap.get(columnName)][currentRi] = (cellType == 's') ? sharedStrings[Integer.valueOf(xsr.getText())] : xsr.getText();
                            }
                        }
                    }
                }

                // If there is a next element go to the next element, else we're done
                if(xsr.hasNext()){ xsr.next(); } else { break; }
            }
        }

        // Return the sheets data as a multi dimentional array
        return output;
    }


    /**
     * @description Method that takes an unzipped xlsx archive as parameter and
     *              returns the sheet data in a multi dimentional array in the form of
     *              data[WorksheetIndex][ColumnIndex][RowIndex]
     * @param  entries A map of path/zip entries from the unzipped XLSX file
     * @return      A list of sheets with a list of rows with a list of cells
     * @note        Faster CPU time, but does not scale very well past 50k cells
     *              Use this method when you work with small sheets only
     */
    global static Object[][][] toArrayDomDoc(Map<String,Compression.ZipEntry> entries){
        return toArrayDomDoc(entries, null);
    }


    /**
     * @description A method to parse the XML using the DOM.Document classes. There is a trade-off in the max size
     *              but this method is significantly faster on the CPU time.
     * @param entries        A map of path/zip entries from the unzipped XLSX file
     * @param selectedSheets A set of sheet indexes that are to be parsed
     * @return      A list of sheets with a list of rows with a list of cells
     * @note        I have optimised this method to being able to handle as much data as possible
     *              this meant that some data is in variable to achieve this. With magical results.
     * @note        Faster CPU time, but does not scale very well past 50k cells
     *              Use this method when you work with small sheets only
     */
    global static Object[][][] toArrayDomDoc(Map<String,Compression.ZipEntry> entries, Set<Integer> selectedSheets){

        // Create a multi-dimentional array for the output
        // Worksheet - Column - Row
        Object[][][] output = new Object[][][]{};

        // Create a list of shared strings
        String[] sharedStrings = getSharedStrings(entries);

        // In order to sort the worksheets, we need to to put them in a list to sort
        String[] worksheetPaths = getWorksheetPaths(entries);

        // Keep track of the largest number of columns. If one sheet is larger than the other
        // we do not do have a duplicate translation for overlapping column names
        // i.e. If sheet 1 has 10 colums and sheet 2 has 20 columns, make sure to only
        // translate the column names from 11 to 20 to save some CPU time
        Integer maxColDimension = 0;

        // Keep track of skipped worksheets
        Integer worksheetSkipCount = 0;

        // Parse each worksheet in it's own array
        for(Integer wi =0, wiMax=worksheetPaths.size(); wi<wiMax; wi++){

            // Skip sheets that are not included in the selectedSheets
            if(selectedSheets != null && !selectedSheets.contains(wi)){
                worksheetSkipCount++;
                continue;
            }

            // Calculate the current worksheet index in case we select specific worksheets only
            Integer currentWi = wi-worksheetSkipCount;

            // Not every row has to be populated, so keep track for the actual row we are in
            Integer currentRi;

            // Add array for the sheet to hold the columns and rows
            output.add(new Object[][]{});

            // Create a DOM document for reading
            Dom.Document doc = new Dom.Document();

            // Load the XML
            doc.load(entries.get(worksheetPaths[wi]).getContent().toString());

            // Extract the dimension
            Dom.XmlNode dimension = doc.getRootElement().getChildElement('dimension', Constant.NS_SPREADSHEET_ML);

            // Get the cell name from the range, this tag is mandatory so should be there
            String cName   = dimension.getAttributeValue('ref', null).split(':')[1];

            // Get the ref element from the token and calculate the value
            String colName = ParseUtil.columnNameFromCellName(cName);

            // Get the number of columsn and rows that are in this worksheet
            Integer numCols = CommonUtil.columnNumberFromColumnName(colName);
            Integer numRows = Integer.valueOf(cName.removeStart(colName));

            // Pre-populate a map with the column name related to the column index i.e. AA => 27, AB =>28 etc
            // Prepopulating makes the parsing go a lot faster as we only have to calculate it once
            if(numCols > maxColDimension){

                // Popualte the colum index (startnumber and the total number of columns)
                ParseUtil.populateColumnIndexMap(maxColDimension, numCols);

                // Update to the new max column dimension
                maxColDimension = numCols;
            }

            // Add empty values for each (potential cell)
            for(Integer ci=0; ci<numCols; ci++){

                // Pre-populate the entire "grid" with empty lists and null values
                output[currentWi].add(new Object[]{});

                // Create a new entry for the row
                for(Integer ri=0; ri<numRows; ri++){
                    output[currentWi][ci].add(null);
                }
            }

            // Get the number of rows for this sheet
            Integer numberOfRows = doc.getRootElement().getChildElement('sheetData',Constant.NS_SPREADSHEET_ML).getChildElements().size();

            // Iterate the rows
            for(Integer ri=0;ri<numberOfRows; ri++){

                // Extracting the row reduces the CPU time by 50%, abnd for some weird reason improves the max heap size
                Dom.XmlNode row = doc.getRootElement().getChildElement('sheetData',Constant.NS_SPREADSHEET_ML).getChildElements()[ri];

                // Calculate the row index
                currentRi = Integer.valueOf(row.getAttribute('r', null)) -1;

                // Get the number of cells in the row
                Integer numberOfCells = row.getChildElements().size();

                // Iterate the cells
                for(Integer ci=0;ci<numberOfCells; ci++){

                    // Extracting the cell reducues the CPU time by another 50% !!
                    Dom.XmlNode cell = row.getChildElements()[ci];

                    String columnName = ParseUtil.columnNameFromCellName(cell.getAttributeValue('r', null));
                    String cellType   = cell.getAttributeValue('t', null);

                    // Define the value in case of an inline string or a shared string
                    Object value = (cellType == 'inlineStr') ?
                                    cell.getChildElement('is',Constant.NS_SPREADSHEET_ML).getChildElement('t', Constant.NS_SPREADSHEET_ML).getText() :
                                    cell.getChildElement('v', Constant.NS_SPREADSHEET_ML).getText()
                    ;

                    // Put the value inside
                    output[currentWi][ParseUtil.columnIndexMap.get(columnName)][currentRi] = (cellType == 's') ? sharedStrings[Integer.valueOf(value)] : value;
                }
            }
        }

        // Return the sheets data as a multi dimentional array
        return output;
    }


    /**
     * @description Method that takes an unzipped xlsx archive as parameter and
     *              returns the sheet data in a multi dimentional array in the form of
     *              data[WorksheetIndex][ColumnIndex][RowIndex]
     * @param  entries A map of path/zip entries from the unzipped XLSX file
     * @return      A list of sheets with a list of rows with a list of cell
     * @note        Use this method for the large sheets
     */
    global static Object[][][] toArrayInverted(Map<String,Compression.ZipEntry> entries){
        return toArrayInverted(entries, null);
    }


    /**
     * @description Method that takes an unzipped xlsx archive as parameter and
     *              returns the sheet data in a multi dimentional array in the form of
     *              data[WorksheetIndex][ColumnIndex][RowIndex]
     * @param  entries        A map of path/zip entries from the unzipped XLSX file
     * @param  selectedSheets A set of integers representing the sheet indexes
     * @return      A list of sheets with a list of rows with a list of cells
     * @note        Use this method for the large sheets
     */
    global static Object[][][] toArrayInverted(Map<String, Compression.ZipEntry> entries, Set<Integer> selectedSheets){

        // Create a multi-dimentional array for the output
        // Worksheet - Column - Row
        Object[][][] output = new Object[][][]{};

        // Create a list of shared strings
        String[] sharedStrings = getSharedStrings(entries);

        // In order to sort the worksheets, we need to to put them in a list to sort
        String[] worksheetPaths = getWorksheetPaths(entries);

        // Keep track of the largest number of columns. If one sheet is larger than the other
        // we do not do have a duplicate translation for overlapping column names
        // i.e. If sheet 1 has 10 colums and sheet 2 has 20 columns, make sure to only
        // translate the column names from 11 to 20 to save some CPU time
        Integer maxColDimension = 0;

        // Keep track of skipped worksheets
        Integer worksheetSkipCount = 0;

        // Parse each worksheet in it's own array
        for(Integer wi =0, wiMax=worksheetPaths.size(); wi<wiMax; wi++){

            // Skip sheets that are not included in the selectedSheets
            if(selectedSheets != null && !selectedSheets.contains(wi)){
                worksheetSkipCount++;
                continue;
            }

            // Calculate the current worksheet index in case we select specific worksheets only
            Integer currentWi = wi-worksheetSkipCount;

            // Add array for the sheet to hold the columns and rows
            output.add(new Object[][]{});

            // Create an XSR for the current worksheet XML file
            XmlStreamReader xsr = new XmlStreamReader(entries.get(worksheetPaths[wi]).getContent().toString());

            // Set the current row to -1 so the indexes match
            Integer currentRi = -1;

            // Variables to hold the type and name of the current open cell element
            String cellType;
            String columnName;

            // Traverse the XML elements in the worksheet
            // I am not using while xsr.hasNext() to save a method call at the end of each item
            while(true) {

                // Start at the beginning of the book and make sure that it is a book
                if (xsr.getEventType() == XmlTag.START_ELEMENT) {
                    switch on xsr.getLocalName() {

                        // First we need the dimention of the worksheet to find out the max size of our rows
                        // The order in the XML is mandatory, so this value should exist, if not the sheet is invalid
                        when 'dimension'{

                            // Get the cell name from the range, this tag is mandatory so should be there
                            String cName   = xsr.getAttributeValueAt(0).split(':')[1];

                            // Get the ref element from the token and calculate the value
                            String colName = ParseUtil.columnNameFromCellName(cName);

                            // Get the number of columsn and rows that are in this worksheet
                            Integer numCols = CommonUtil.columnNumberFromColumnName(colName);
                            Integer numRows = Integer.valueOf(cName.removeStart(colName));

                            // Pre-populate a map with the column name related to the column index i.e. AA => 27, AB =>28 etc
                            // Prepopulating makes the parsing go a lot faster as we only have to calculate it once
                            if(numCols > maxColDimension){

                                // Popualte the colum index (startnumber and the total number of columns)
                                ParseUtil.populateColumnIndexMap(maxColDimension, numCols);

                                // Update to the new max column dimension
                                maxColDimension = numCols;
                            }

                            // Add empty values for each (potential cell)
                            for(Integer ri=0; ri<numRows; ri++){

                                // Pre-populate the entire "grid" with empty lists and null values
                                output[currentWi].add(new Object[]{});

                                // Create a new entry for the row
                                for(Integer ci=0; ci<numCols; ci++){
                                    output[currentWi][ri].add(null);
                                }
                            }
                        }

                        // The row element, keep track of what row we're currently in
                        // We can have a situation where we have row 1 and 10 populated, but nothing in between
                        when 'row'{
                            currentRi = Integer.valueOf(xsr.getAttributeValue(null, 'r'))-1;
                        }

                        // The cell element
                        when 'c' {
                            columnName = ParseUtil.columnNameFromCellName(xsr.getAttributeValue(null, 'r'));
                            cellType   = xsr.getAttributeValue(null, 't');
                        }

                        // Value tags inside the cell element
                        when 'v','t'{

                            // Skip to the text element
                            xsr.next();

                            // Validate we're in a text element
                            if(xsr.getEventType() == XmlTag.CHARACTERS){

                                // Put the cell value in array at location wi.ci.ri
                                // If there is a "sharedStrings index", get the shared strings value at that index
                                // else the value in the value in the XSR
                                output[currentWi][currentRi][ParseUtil.columnIndexMap.get(columnName)] = (cellType == 's') ? sharedStrings[Integer.valueOf(xsr.getText())] : xsr.getText();
                            }
                        }
                    }
                }

                // If there is a next element go to the next element, else we're done
                if(xsr.hasNext()){ xsr.next(); } else { break; }
            }
        }

        // Return the sheets data as a multi dimentional array
        return output;
    }

    /** **************************************************************************************************** **
     **                                          GLOBAL MAP METHODS                                          **
     ** **************************************************************************************************** **/
    /**
     * @description Method that takes an unzipped xlsx archive as parameter and
     *              returns the sheet data as a list of maps where the cellName is
     *              used as the map key
     *              data[WorksheetIndex].get(cellName)
     * @param  entries        A map of path/zip entries from the unzipped XLSX file
     * @return      A list of Maps where each map represents a worksheet, each key the cellname
     *              and the map value the cell value
     */
    global static List<Map<String,Object>> toMap(Map<String,Compression.ZipEntry> entries){
        return toMap(entries, null);
    }


    /**
     * @description Method that takes an unzipped xlsx archive as parameter and
     *              returns the sheet data as a list of maps where the cellName is
     *              used as the map key
     *              data[WorksheetIndex].get(cellName)
     * @param entries        A map of path/zip entries from the unzipped XLSX file
     * @param selectedSheets A set of sheet indexes that are to be parsed
     * @return      A list of Maps where each map represents a worksheet, each key the cellname
     *              and the map value the cell value
     */
    global static List<Map<String,Object>> toMap(Map<String,Compression.ZipEntry> entries, Set<Integer> selectedSheets){

        // List of maps to hold the output
        List<Map<String,Object>> output = new List<Map<String,Object>>();

        // Create a list of shared strings
        String[] sharedStrings  = getSharedStrings( entries);

        // In order to sort the worksheets, we need to to put them in a list to sort
        String[] worksheetPaths = getWorksheetPaths(entries);

        // Keep track of skipped worksheets
        Integer worksheetSkipCount = 0;

        // Parse each worksheet in it's own array
        for(Integer wi=0, wiMax=worksheetPaths.size(); wi<wiMax; wi++){

            // Skip sheets that are not included in the selectedSheets
            if(selectedSheets != null && !selectedSheets.contains(wi)){
                worksheetSkipCount++;
                continue;
            }

            // Calculate the current worksheet index in case we select specific worksheets only
            Integer currentWi = wi-worksheetSkipCount;

            // Create a map for this worksheet
            output.add(new Map<String,Object>());

            // Create an XSR for the current worksheet XML file
            XmlStreamReader xsr = new XmlStreamReader(entries.get(worksheetPaths[wi]).getContent().toString());

            // Variables to hold the type and name of the current open cell element
            String cellName;
            String cellType;

            // Traverse the XML elements in the worksheet
            while(true) {

                // Start at the beginning of the book and make sure that it is a book
                if (xsr.getEventType() == XmlTag.START_ELEMENT) {
                    switch on xsr.getLocalName() {

                        // The cell element
                        when 'c' {
                            cellName = xsr.getAttributeValue(null, 'r');
                            cellType = xsr.getAttributeValue(null, 't');
                        }

                        // Value tags inside the cell element
                        when 'v','t'{

                            // Skip to the text element
                            xsr.next();

                            // Validate we're in a text element
                            if(xsr.getEventType() == XmlTag.CHARACTERS){

                                // Put the value in wi.ri.ci, if shared strings get the shared strings value
                                // else the value in the value in the XSR
                                output[currentWi].put(
                                    cellName,
                                    (cellType == 's') ? sharedStrings[Integer.valueOf(xsr.getText())] : xsr.getText()
                                );
                            }
                        }
                    }
                }

                // If there is a next element go to the next element, else we're done
                if(xsr.hasNext()){ xsr.next(); }else{ break; }
            }
        }
        // Return the fully populated output
        return output;
    }


    /**
     * @description Method that takes an unzipped xlsx archive as parameter and
     *              returns the sheet data as a list of maps where the cellName is
     *              used as the map key
     *              data[WorksheetIndex].get(cellName)
     * @param  entries        A map of path/zip entries from the unzipped XLSX file
     * @param  selectedSheets A set of integers representing the sheet indexes
     * @return      A list of Maps where each map represents a worksheet, each key the cellname
     *              and the map value the cell value
     * @note        Faster CPU time, but does not scale very well past 50k cells
     *              Use this method when you work with small sheets only
     */
    global static List<Map<String,Object>> toMapDomDoc(Map<String,Compression.ZipEntry> entries){
        return toMapDomDoc(entries,null);
    }


    /**
     * @description Method that takes an unzipped xlsx archive as parameter and
     *              returns the sheet data as a list of maps where the cellName is
     *              used as the map key
     *              data[WorksheetIndex].get(cellName)
     * @param entries        A map of path/zip entries from the unzipped XLSX file
     * @param selectedSheets A set of sheet indexes that are to be parsed
     * @return      A list of Maps where each map represents a worksheet, each key the cellname
     *              and the map value the cell value
     * @note        Faster CPU time, but does not scale very well past 50k cells
     *              Use this method when you work with small sheets only
     */
    global static List<Map<String,Object>> toMapDomDoc(Map<String,Compression.ZipEntry> entries, Set<Integer> selectedSheets){

        // List of maps to hold the output
        List<Map<String,Object>> output = new List<Map<String,Object>>();

        // Create a list of shared strings
        String[] sharedStrings  = getSharedStrings( entries);

        // In order to sort the worksheets, we need to to put them in a list to sort
        String[] worksheetPaths = getWorksheetPaths(entries);

        // Keep track of skipped worksheets
        Integer worksheetSkipCount = 0;

        // Parse each worksheet in it's own array
        for(Integer wi=0, wiMax=worksheetPaths.size(); wi<wiMax; wi++){

            // Skip sheets that are not included in the selectedSheets
            if(selectedSheets != null && !selectedSheets.contains(wi)){
                worksheetSkipCount++;
                continue;
            }

            // Calculate the current worksheet index in case we select specific worksheets only
            Integer currentWi = wi-worksheetSkipCount;

            // Create a map for this worksheet
            output.add(new Map<String,Object>());

            // Create a DOM document for reading
            Dom.Document doc = new Dom.Document();

            // Load the XML
            doc.load(entries.get(worksheetPaths[wi]).getContent().toString());

            // Get the number of rows for this sheet
            Integer numberOfRows = doc.getRootElement().getChildElement('sheetData',Constant.NS_SPREADSHEET_ML).getChildElements().size();

            // Iterate the rows
            for(Integer ri=0;ri<numberOfRows; ri++){

                // Extracting the row reduces the CPU time by 50%, abnd for some weird reason improves the max heap size
                Dom.XmlNode row = doc.getRootElement().getChildElement('sheetData',Constant.NS_SPREADSHEET_ML).getChildElements()[ri];

                // Get the number of cells in the row
                Integer numberOfCells = row.getChildElements().size();

                // Iterate the cells
                for(Integer ci=0;ci<numberOfCells; ci++){

                    // Extracting the cell reducues the CPU time by another 50% !!
                    Dom.XmlNode cell = row.getChildElements()[ci];

                    // Get the cell name and the cell type
                    String cellName = cell.getAttributeValue('r', null);
                    String cellType = cell.getAttributeValue('t', null);

                    // Define the value in case of an inline string or a shared string
                    Object value = (cellType == 'inlineStr') ?
                                    cell.getChildElement('is',Constant.NS_SPREADSHEET_ML).getChildElement('t', Constant.NS_SPREADSHEET_ML).getText() :
                                    cell.getChildElement('v', Constant.NS_SPREADSHEET_ML).getText()
                    ;

                    // Put the value in wi.ri.ci, if shared strings get the shared strings value
                    // else the value in the value in the XSR
                    output[currentWi].put(
                        cellName,
                        (cellType == 's') ? sharedStrings[Integer.valueOf(value)] : value
                    );
                }
            }
        }
        // Return the fully populated output
        return output;
    }


    /** **************************************************************************************************** **
     **                              PRIVATE SUPPORT METHODS (CLASS SPECIFIC)                                **
     ** **************************************************************************************************** **/
    /**
     * @description This method takes a Map of file paths and returns an array of shared strings.
     * @param entries A Map of file paths.
     * @return      String[] An array of shared strings.
     */
    private static String[] getSharedStrings(Map<String,Compression.ZipEntry> entries){

        String[] sharedStrings = new String[]{};

        // Check the shared strings file exists
        if(entries.containsKey(Constant.ZP_PATH_SHRD_STRS)){

            // Create an XSR for the shared strings XML file
            XmlStreamReader xsr = new XmlStreamReader(entries.get(Constant.ZP_PATH_SHRD_STRS).getContent().toString());

            // Indicator if there are any more XML elements left to traverse
            boolean isSafeToGetNextXmlElement = true;

            // Traverse the XML elements in the worksheet
            while(isSafeToGetNextXmlElement){

                // Get all the string values and add them to the shared stings list
                if (xsr.getEventType() == XmlTag.START_ELEMENT && xsr.getLocalName() == 't') {
                    xsr.next();

                    // Add the value to the shared strings, if the tag is empty add a blank value
                    sharedStrings.add(xsr.getEventType() == XmlTag.CHARACTERS ? xsr.getText() : '');
                }

                // If there is a next element go to the next element, else we're done
                if(xsr.hasNext()){ xsr.next(); } else { break; }
            }
        }

        // Return the array of shared strings
        return sharedStrings;
    }


    /**
     * @description This method takes a Map of file paths and returns an array of worksheet paths.
     * @param entries A Map of file paths.
     * @return      String[] An array of worksheet paths.
    */
    private static String[] getWorksheetPaths(Map<String,Compression.ZipEntry> entries){

        // List to hold the worksheet paths
        String[] worksheetPaths = new String[]{};

        // Iterate through all the files in the workbook folder
        for(String path : entries.keySet()){
            if(path.startsWith(Constant.FOLDER_XL_WORKSHEET)){
                worksheetPaths.add(path);
            }
        }
        // Sort the worksheet paths so they will be in the correct order
        worksheetPaths.sort();

        // Return the sort list of worksheet paths
        return worksheetPaths;
    }


    /** **************************************************************************************************** **
     **                                    PRIVATE DATA STRUCTURE CLASSES                                    **
     ** **************************************************************************************************** **/
    /**
     * @description Class representing a sheet index xml element
     *              Used to create a sortable list based on the index
     */
    private class SheetIndex implements Comparable{

        // Public variables
        private Integer index;
        private String  name;

        /**
         * @description Main constructor
         * @param index The sheet index
         * @param name  The sheet name
         */
        public SheetIndex(Integer index, String name){
            this.index = index;
            this.name  = name ;
        }

        /**
         * @description Method for comparable interface
         * @param compareTo The object to compare against
         * @return      1, 0 or -1
         */
        public Integer compareTo(Object compareTo){
            return (this.index > ((SheetIndex) compareTo).index) ? 1 : (this.index == ((SheetIndex) compareTo).index) ? 0 : -1;
        }
    }
}
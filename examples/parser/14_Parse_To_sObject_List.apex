/**
 * @author         Justus van den Berg (jfwberg@gmail.com)
 * @date           January 2025
 * @copyright      (c) 2025 Justus van den Berg
 * @license        MIT (See LICENSE file in the project root)
 * @description    Examples on how to parse an XLSX file stored as a document
 * @examples       Parse to Multi-Dimensional Array and convert
 *                 the result to a list of SObjects
 * @note           This method uses the toArrayInverted() method to have record
 *                 rows before columns instead of the other way around
 * @note           - Worksheets need to have the sObject API Name
 *                 - The first line of the worksheet requires the API fields in the header
 *                 - No field validation like read-only or formula fields is applied
 */
/** ****************************************************************************** **
 *     Parse to Multi-Dimensional Array and convert to a list of sObject Lists
 ** ****************************************************************************** **/
try{
    // If there are any errors or warnings they will be put in this list
    String[] warnings = new String[]{};

    // Create a new list of sObject lists for the output of each worksheet
    SObject[][] sObjectLists = new SObject[][]{};
    
    // Unzip the whole xlsx file, make sure the XLSX file only contains existing sObjects
    Map<String,Compression.ZipEntry> entries = new Compression.ZipReader(
        [SELECT VersionData FROM ContentVersion WHERE ContentDocumentId = '069Qz000003N27SIAS' AND IsLatest = true LIMIT 1]?.VersionData
    ).getEntriesMap();

    // Get an index for all the worksheets, because we need the worksheet name as sObject API name
    Map<String,Integer> worksheetNameIndexMap =  xlsx.Parse.toWorksheetNameIndexMap(entries); 
    
    // Create a list with the worksheet names so we can reference it using the index
    String[] worksheetNameList = new List<String>(worksheetNameIndexMap.keySet());

    // Keep track of the current worksheet
    Integer wi = -1;

    // As data array
    for(Object[][] worksheet : xlsx.Parse.toArrayInverted(entries)){
        
        // Set the WI
        wi++;

        // Create a list of sObjects for the current worksheet
        SObject[] sObjectList = (SObject[]) Type.forName( 'List<'+worksheetNameList[wi]+'>' ).newInstance();

        // Get the fields for the current sObject
        Map<String, Schema.SObjectField> fieldMap = sObjectList.getSObjectType().getDescribe().fields.getMap();

        // Get a list of field names from the first row
        String[] fieldNames = new String[]{};

        // Convert the field name row to a list of strings
        for(Integer i=0, max=worksheet[0].size();i<max;i++){
            fieldNames.add((String)worksheet[0][i]);
        }

        // Iterate the rows
        for(Integer ri=1, max=worksheet.size(); ri<max; ri++){
            
            // Create a new sObject
            SObject sObj = (SObject) Type.forName( worksheetNameList[wi] ).newInstance();

            // Populate the sObject with the values from the current row
            for(Integer ci=0, cMax=worksheet[0].size();ci<cMax;ci++){
                
                if(!fieldMap.containsKey( fieldNames[ci] )){
                    warnings.add(worksheetNameList[wi] + '.' + fieldNames[ci] + ' does not exist and is skipped');
                    continue;
                }

                if(!fieldMap.get(fieldNames[ci]).getDescribe().isCreateable() || !fieldMap.get(fieldNames[ci]).getDescribe().isUpdateable()){
                    warnings.add(worksheetNameList[wi] + '.' + fieldNames[ci] + ' is not creteable or updateable and skipped');
                    continue;
                }
                
                System.debug(fieldMap.get(fieldNames[ci]).getDescribe().getType() + fieldNames[ci] + ' --- ' + worksheet[ri][ci]);
                
                sObj.put(
                    fieldNames[ci],
                    castObjectFieldData(
                        fieldMap.get(fieldNames[ci]).getDescribe().getType(),
                        worksheet[ri][ci]
                    )
                );
            }

            // Add the sObject to the list
            sObjectList.add( sObj );
        }

        // Add the sObject list to the output
        sObjectLists.add(sObjectList);
    }

    // Output the results of what we have done
    for(SObject[] sObjectList : sObjectLists){
        System.debug('sObjectType: "' + sObjectList.getSObjectType() + '" contains ' + sObjectList.size() + ' records');
    }

    // Output any warnings on the import
    System.debug(warnings);
    
// Catch a Zip Specific Exception
}catch(Compression.ZipException e){
    System.debug('ZIP EXCEPTION: ' + e.getMessage());

// Catch an XLSX Parser Specific Exception
}catch(xlsx.ParseException e){
    System.debug('PARSE EXCEPTION: ' + e.getMessage());

// Catch and handle any other type of Exception
}catch(Exception e){
    System.debug('UNKNOWN EXCEPTION: ' + e.getMessage() + e.getStackTraceString());
}


public static Object castObjectFieldData(Schema.DisplayType fieldType, Object fieldValue) {
    // create a switch for each dispaytype
    switch on fieldType {
        when BOOLEAN {
            return Boolean.valueOf(fieldValue);
        }
        when DOUBLE, CURRENCY, PERCENT {
            return Double.valueOf(fieldValue);
        }
        when DATE {
            return Date.valueOf(String.valueOf(fieldValue));
        }
        when DATETIME {
            return Datetime.valueOf(fieldValue);
        }
        when INTEGER {
            return Integer.valueOf(fieldValue);
        }
        when STRING, TEXTAREA, URL, PHONE, EMAIL {
            return String.valueOf(fieldValue);
        }
        when else {
            return fieldValue;
        }
    }
}
/**
 * @author         Justus van den Berg (jfwberg@gmail.com)
 * @date           January 2025
 * @copyright      (c) 2025 Justus van den Berg
 * @license        MIT (See LICENSE file in the project root)
 * @description    Test class for the Parse and Build classes
 * @tests          Parse
 *                 Build
 */
@IsTest
@SuppressWarnings('PMD.OneDeclarationPerLine')
private with sharing class ParseTest {

    /** **************************************************************************************************** **
     **                                       GLOBAL WORKSHEET METHODS                                       **
     ** **************************************************************************************************** **/
    /**
     * @description Test Method for the worksheet name to index logic
     */
    @IsTest
    private static void testToWorksheetNameIndexMap(){

        // Create a basic file with 2 worksheets
        Builder b = new Builder();
        b.addWorksheet('Worksheet 01');
        b.addWorksheet('Worksheet 02');
        
        // Build the archive as document and content version
        Blob documentData       = Build.asDocument(b).body;
        Blob contentVersionData = Build.asContentVersion(b).VersionData;

        Test.startTest();
        Map<String,Integer> indexMap1 = Parse.toWorksheetNameIndexMap(new Compression.ZipReader(documentData).getEntriesMap());
        Map<String,Integer> indexMap2 = Parse.toWorksheetNameIndexMap(new Compression.ZipReader(contentVersionData).getEntriesMap());
        Test.stopTest();

        // Assert the indexes are correct
        Assert.areEqual(0, indexMap1.get('Worksheet 01'), 'Unexpected sheet index');
        Assert.areEqual(1, indexMap1.get('Worksheet 02'), 'Unexpected sheet index');
        Assert.areEqual(0, indexMap2.get('Worksheet 01'), 'Unexpected sheet index');
        Assert.areEqual(1, indexMap2.get('Worksheet 02'), 'Unexpected sheet index');
    }

    
    @IsTest
    private static void testToWorksheetNameIndexMapException(){
        
        // Create invalid archive
        Compression.ZipWriter zw = new Compression.ZipWriter();
        zw.addEntry('invalid.txt', Blob.valueOf('Some text'));
        Blob archive = zw.getArchive();
        
        // Run the test
        try{
            Parse.toWorksheetNameIndexMap(new Compression.ZipReader(archive).getEntriesMap());
            throw new TestException(TestUtil.EXCEPTION_SAFEGUARD_MESSAGE);
        } catch(ParseException e){
            Assert.AreEqual(
                String.format(Constant.MSG_INVALID_ZIP_ENTRY, new String[]{Constant.ZP_PATH_WORKBOOK}),
                e.getMessage(),
                TestUtil.EXCEPTION_ASSERTION_MESSAGE
            );
        }
    }


    /** **************************************************************************************************** **
     **                                         GLOBAL ARRAY METHODS                                         **
     ** **************************************************************************************************** **/
    /**
     * @description Test Method for the toArray logic
     */
    @IsTest
    private static void testToArray() {
        
        // Create a new builder
        Builder b = new Builder();
        
        // Create the test data
        populateWorksheets(b);
        
        // Execute the tests
        Test.startTest();
        Object[][][] documentResult                 = Parse.toArray(
                                                        new Compression.ZipReader(Build.asDocument(b).body).getEntriesMap()
                                                    );

        Object[][][] contentVersionResult           = Parse.toArray(
                                                        new Compression.ZipReader(Build.asContentVersion(b).VersionData).getEntriesMap(),
                                                        new Set<Integer>{0,2}
                                                    );

        Object[][][] documentResultDomDoc           = Parse.toArrayDomDoc(
                                                        new Compression.ZipReader(Build.asDocument(b).body).getEntriesMap()
                                                    );

        Object[][][] contentVersionResultDomDoc     = Parse.toArrayDomDoc(
                                                        new Compression.ZipReader(Build.asContentVersion(b).VersionData).getEntriesMap(),
                                                        new Set<Integer>{0,2}
                                                    );
        
        Object[][][] documentResultIv               = Parse.toArrayInverted(
                                                        new Compression.ZipReader(Build.asDocument(b).body).getEntriesMap()
                                                    );

        Object[][][] contentVersionResultIv         = Parse.toArrayInverted(
                                                        new Compression.ZipReader(Build.asContentVersion(b).VersionData).getEntriesMap(),
                                                        new Set<Integer>{0,2}
                                                    );

        Object[][][] documentResultDomDocIv         = Parse.toArrayInvertedDomDoc(
                                                        new Compression.ZipReader(Build.asDocument(b).body).getEntriesMap()
                                                    );

        Object[][][] contentVersionResultDomDocIv   = Parse.toArrayInvertedDomDoc(
                                                        new Compression.ZipReader(Build.asContentVersion(b).VersionData).getEntriesMap(),
                                                        new Set<Integer>{0,2}
                                                    );
        Test.stopTest();

        // Validate we have the correct number of worksheets
        Assert.areEqual(3,documentResult.size(),                'Unexpected number of worksheets');
        Assert.areEqual(2,contentVersionResult.size(),          'Unexpected number of worksheets');
        Assert.areEqual(3,documentResultDomDoc.size(),          'Unexpected number of worksheets');
        Assert.areEqual(2,contentVersionResultDomDoc.size(),    'Unexpected number of worksheets');
        Assert.areEqual(3,documentResultIv.size(),              'Unexpected number of worksheets');
        Assert.areEqual(2,contentVersionResultIv.size(),        'Unexpected number of worksheets');
        Assert.areEqual(3,documentResultDomDocIv.size(),        'Unexpected number of worksheets');
        Assert.areEqual(2,contentVersionResultDomDocIv.size(),  'Unexpected number of worksheets');

        // Iterate all worksheets in all results and validate their output
        for (Integer i = 0, max=documentResult.size(); i < max; i++) {
            validateCellValuesInArray(documentResult[i]);
        }

        // Iterate all worksheets in all results and validate their output
        for (Integer i = 0, max=contentVersionResult.size(); i < max; i++) {
            validateCellValuesInArray(contentVersionResult[i]);
        }

        // Iterate all worksheets in all results and validate their output
        for (Integer i = 0, max=documentResultDomDoc.size(); i < max; i++) {
            validateCellValuesInArray(documentResultDomDoc[i]);
        }

        // Iterate all worksheets in all results and validate their output
        for (Integer i = 0, max=contentVersionResultDomDoc.size(); i < max; i++) {
            validateCellValuesInArray(contentVersionResultDomDoc[i]);
        }

        // Iterate all worksheets in all results and validate their output
        for (Integer i = 0, max=documentResultIv.size(); i < max; i++) {
            validateCellValuesInArray(documentResultIv[i]);
        }

        // Iterate all worksheets in all results and validate their output
        for (Integer i = 0, max=contentVersionResultIv.size(); i < max; i++) {
            validateCellValuesInArray(contentVersionResultIv[i]);
        }

        // Iterate all worksheets in all results and validate their output
        for (Integer i = 0, max=documentResultDomDocIv.size(); i < max; i++) {
            validateCellValuesInArray(documentResultDomDocIv[i]);
        }

        // Iterate all worksheets in all results and validate their output
        for (Integer i = 0, max=contentVersionResultDomDocIv.size(); i < max; i++) {
            validateCellValuesInArray(contentVersionResultDomDocIv[i]);
        } 
    }


    /** **************************************************************************************************** **
     **                                          GLOBAL MAP METHODS                                          **
     ** **************************************************************************************************** **/
    /**
     * @description Test Method for the toMap logic
     */
    @IsTest
    private static void testToMap() {
        
        // Create a new builder
        Builder b = new Builder();
        
        // Test that the keywords get cropped at 255 characters
        for (Integer i = 0; i < 30; i++) {
            b.addKeyword('1234567890');
        }
        
        // Create the test data
        populateWorksheets(b);
        
        // Execute the tests
        Test.startTest();
        List<Map<String,Object>> documentResult             = Parse.toMap(
                                                                new Compression.ZipReader(Build.asDocument(b).body).getEntriesMap()
                                                            );

        List<Map<String,Object>> contentVersionResult       = Parse.toMap(
                                                                new Compression.ZipReader(Build.asContentVersion(b).VersionData).getEntriesMap(),
                                                                new Set<Integer>{0,2}
                                                            );

        List<Map<String,Object>> documentResultDomDoc       = Parse.toMapDomDoc(
                                                                new Compression.ZipReader(Build.asDocument(b).body).getEntriesMap()
                                                            );

        List<Map<String,Object>> contentVersionResultDomDoc = Parse.toMapDomDoc(
                                                                new Compression.ZipReader(Build.asContentVersion(b).VersionData).getEntriesMap(),
                                                                new Set<Integer>{0,2}
                                                            );
        Test.stopTest();

        // Validate we have the correct number of worksheets
        Assert.areEqual(3,documentResult.size(),             'Unexpected number of worksheets');
        Assert.areEqual(2,contentVersionResult.size(),       'Unexpected number of worksheets');
        Assert.areEqual(3,documentResultDomDoc.size(),       'Unexpected number of worksheets');
        Assert.areEqual(2,contentVersionResultDomDoc.size(), 'Unexpected number of worksheets');

        // Iterate all worksheets in all results and validate their output
        for (Integer i = 0, max=documentResult.size(); i < max; i++) {
            validateCellValuesInMap(documentResult[i]);
        }

        // Iterate all worksheets in all results and validate their output
        for (Integer i = 0, max=contentVersionResult.size(); i < max; i++) {
            validateCellValuesInMap(contentVersionResult[i]);
        }

        // Iterate all worksheets in all results and validate their output
        for (Integer i = 0, max=documentResultDomDoc.size(); i < max; i++) {
            validateCellValuesInMap(documentResultDomDoc[i]);
        }

        // Iterate all worksheets in all results and validate their output
        for (Integer i = 0, max=contentVersionResultDomDoc.size(); i < max; i++) {
            validateCellValuesInMap(contentVersionResultDomDoc[i]);
        } 
    }


    /** **************************************************************************************************** **
     **                                             OTHER METHODS                                            **
     ** **************************************************************************************************** **/
    /**
     * @description Test Method for the inline strings logic
     */
    @IsTest
    private static void testInlineStrings() {
        
        // Create a new builder
        Builder b = new Builder();
        
        // Use inline strings
        b.setUseSharedStrings(false);
        b.addTextCell(0, 0, 0, 'Lalalalala');
        
        // Execute the tests
        Test.startTest();

        Object[][][]             arrayResult    = Parse.toArrayDomDoc(
                                                    new Compression.ZipReader(Build.asDocument(b).body).getEntriesMap()
                                                );

        Object[][][]             arrayResultIv  = Parse.toArrayInvertedDomDoc(
                                                    new Compression.ZipReader(Build.asDocument(b).body).getEntriesMap()
                                                );

        List<Map<String,Object>> mapResult      = Parse.toMapDomDoc(
                                                    new Compression.ZipReader(Build.asDocument(b).body).getEntriesMap()
                                                );
        Test.stopTest();

        // Validate we have the correct number of worksheets
        Assert.areEqual(1, arrayResult.size(),   'Unexpected number of worksheets');
        Assert.areEqual(1, arrayResultIv.size(), 'Unexpected number of worksheets');
        Assert.areEqual(1, mapResult.size(),     'Unexpected number of worksheets');
        
        // Validate cell values
        Assert.areEqual('Lalalalala',   String.valueOf( arrayResult[0][0][0]),   'Unexpected Cell value');
        Assert.areEqual('Lalalalala',   String.valueOf( mapResult[0].get('A1')), 'Unexpected Cell value');
    }


    /** **************************************************************************************************** **
     **                              PRIVATE SUPPORT METHODS (CLASS SPECIFIC)                                **
     ** **************************************************************************************************** **/
    /**
     * @description Method to created worksheet test data
     * @param b The XLSX builder class instance
     */
    private static void populateWorksheets(Builder b){
        for (Integer i = 0; i < 3; i++) {
            b.addTextCell(   i, 0, 0, 'A1'     );
            b.addTextCell(   i, 1, 1, 'B2',   0);
            b.addNumberCell( i, 2, 2, 123.45   );
            b.addNumberCell( i, 3, 3, 678.90, 0);
            b.addNumberCell( i, 4, 4, 100      );
            b.addNumberCell( i, 5, 5, 200,    0);
            b.addBooleanCell(i, 6, 6, true     );
            b.addBooleanCell(i, 7, 7, false,  0);
            b.addFormulaCell(i, 8, 8, 19,     '=SUM(9,10)'   );
            b.addFormulaCell(i, 9, 9, 20,     '=SUM(10,10)',0);
        }
    }


    /**
     * @description method to validate cell values for multiple workheets
     * @param worksheet The rows and cells array for a specific worksheet
     */
    private static void validateCellValuesInArray(Object[][] worksheet){
        
        // Validate the number of values
        Assert.areEqual(10, worksheet.size(),'Unexpected number of rows in worksheet');

        // Validate each individual values
        Assert.areEqual('A1',   String.valueOf( worksheet[0][0]), 'Unexpected Cell value');
        Assert.areEqual('B2',   String.valueOf( worksheet[1][1]), 'Unexpected Cell value');
        Assert.areEqual(123.45, Double.valueOf( worksheet[2][2]), 'Unexpected Cell value');
        Assert.areEqual(678.90, Double.valueOf( worksheet[3][3]), 'Unexpected Cell value');
        Assert.areEqual(100,    Integer.valueOf(worksheet[4][4]), 'Unexpected Cell value');
        Assert.areEqual(200,    Integer.valueOf(worksheet[5][5]), 'Unexpected Cell value');
        Assert.areEqual(true,   Boolean.valueOf(worksheet[6][6]), 'Unexpected Cell value');
        Assert.areEqual(false,  Boolean.valueOf(worksheet[7][7]), 'Unexpected Cell value');
        Assert.areEqual(19,     Integer.valueOf(worksheet[8][8]), 'Unexpected Cell value');
        Assert.areEqual(20,     Integer.valueOf(worksheet[9][9]), 'Unexpected Cell value');
    }


    /**
     * @description method to validate cell values for multiple workheets
     * @param worksheet The rows and cells array for a specific worksheet
     */
    private static void validateCellValuesInMap(Map<String,Object> worksheet){
        
        // Validate the number of values
        Assert.areEqual(10, worksheet.size(),'Unexpected number of rows in worksheet');

        // Validate each individual values
        Assert.areEqual('A1',   String.valueOf( worksheet.get('A1' )), 'Unexpected Cell value');
        Assert.areEqual('B2',   String.valueOf( worksheet.get('B2' )), 'Unexpected Cell value');
        Assert.areEqual(123.45, Double.valueOf( worksheet.get('C3' )), 'Unexpected Cell value');
        Assert.areEqual(678.90, Double.valueOf( worksheet.get('D4' )), 'Unexpected Cell value');
        Assert.areEqual(100,    Integer.valueOf(worksheet.get('E5' )), 'Unexpected Cell value');
        Assert.areEqual(200,    Integer.valueOf(worksheet.get('F6' )), 'Unexpected Cell value');
        Assert.areEqual(true,   Boolean.valueOf(worksheet.get('G7' )), 'Unexpected Cell value');
        Assert.areEqual(false,  Boolean.valueOf(worksheet.get('H8' )), 'Unexpected Cell value');
        Assert.areEqual(19,     Integer.valueOf(worksheet.get('I9' )), 'Unexpected Cell value');
        Assert.areEqual(20,     Integer.valueOf(worksheet.get('J10')), 'Unexpected Cell value');
    }
}
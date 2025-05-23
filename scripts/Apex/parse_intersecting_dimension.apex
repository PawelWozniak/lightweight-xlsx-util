/**
 * @description Multi-Dimensional Array Testing Script to validate all array method
 *              variations return the same number of columns, rows and cells
 */
// As data array
Object[][][] xlsxDataArray = xlsx.Parse.toArray(
    new Compression.ZipReader(
        [SELECT body FROM Document WHERE Id = '015Pt0000055Dub' LIMIT 1]?.Body
    ).getEntriesMap()
);

// As inverted data array
Object[][][] xlsxDataArrayInverted = xlsx.Parse.toArrayInverted(
    new Compression.ZipReader(
        [SELECT body FROM Document WHERE Id = '015Pt0000055Dub' LIMIT 1]?.Body
    ).getEntriesMap()
);

// As data array DomDoc
Object[][][] xlsxDataArrayDomDoc = xlsx.Parse.toArrayDomDoc(
    new Compression.ZipReader(
        [SELECT body FROM Document WHERE Id = '015Pt0000055Dub' LIMIT 1]?.Body
    ).getEntriesMap()
);

// As inverted data array DomDoc
Object[][][] xlsxDataArrayInvertedDomDoc = xlsx.Parse.toArrayInvertedDomDoc(
    new Compression.ZipReader(
        [SELECT body FROM Document WHERE Id = '015Pt0000055Dub' LIMIT 1]?.Body
    ).getEntriesMap()
);

// Process the arrays and extract the different counts
List<Map<String,Integer>> aResult    = processArray(xlsxDataArray,               false );
List<Map<String,Integer>> aiResult   = processArray(xlsxDataArrayInverted,       true  );
List<Map<String,Integer>> addResult  = processArray(xlsxDataArrayDomDoc,         false );
List<Map<String,Integer>> addiResult = processArray(xlsxDataArrayInvertedDomDoc, true  );

// Should be true, as all methods should return exactly the same value for each parse method
Boolean allResultsAreEqual = aResult == aiResult && aResult == addResult && aResult == addiResult;

// Output the number of cells
System.debug(JSON.serializePretty(aResult)   );
System.debug(JSON.serializePretty(aiResult)  );
System.debug(JSON.serializePretty(addResult) );
System.debug(JSON.serializePretty(addiResult));

// Output the assertion result
System.debug('All results are equal: ' + allResultsAreEqual);



/**
 * @description Process the arrays and extract the different counts
 * @param Object[][][]  xlsxDataArray The array to process
 * @param inverted      Indicator if rows come before columns
 */
List<Map<String,Integer>> processArray(Object[][][] xlsxDataArray, Boolean inverted){
    
    // Create the output info for each worksheet
    List<Map<String,Integer>> dataInfo = new List<Map<String,Integer>>();

    // Loop through the worksheets, columns and rows to touch each cell
    for(Integer wi=0, wiMax=xlsxDataArray.size(); wi<wiMax; wi++){
        
        // Create new info point
        dataInfo.add(new Map<String,Integer>());
        
        // Reset counters
        Integer colCounter  = 0;
        Integer rowCounter  = 0;
        Integer cellCounter = 0;

        // Invert if required
        if(inverted){
            for(Integer ri=0, riMax=xlsxDataArray[wi].size(); ri<riMax; ri++){
                rowCounter++;
                for(Integer ci=0, ciMax=xlsxDataArray[wi][ri].size(); ci<ciMax; ci++){
                    if(ri==0){colCounter++;}
                    cellCounter++;
                }
            }
        }else{
            for(Integer ci=0, ciMax=xlsxDataArray[wi].size(); ci<ciMax; ci++){
                colCounter++;
                for(Integer ri=0, riMax=xlsxDataArray[wi][ci].size(); ri<riMax; ri++){
                    if(ci==0){rowCounter++;}
                    cellCounter++;
                }
            }
        }

        // Add worksheet specific info to the result output
        dataInfo[wi].put('columnCount', colCounter);
        dataInfo[wi].put('rowCount', rowCounter);
        dataInfo[wi].put('cellCount', cellCounter);
    }

    // Return the data
    return dataInfo;
}
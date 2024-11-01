/**
 * @author         Justus van den Berg (jfwberg@gmail.com)
 * @date           November 2024
 * @copyright      (c) 2023 Justus van den Berg
 * @license        MIT (See LICENSE file in the project root)
 * @description    Class containing utilities for the Parsing of XLSX files
 * @false-positive AvoidGlobalModifier      This is a utility that is designed to be called from a
 *                                          managed package. It can be used for a user's own custom
 *                                          Implementation, so global is the way to open up this
 *                                          utility for global use.
 */
public with sharing class ParseUtil {

    // Map that holds the column name as a key and it's matching index as a value
    // This is for optimising performance by caching instead of repeat calculation
    public static Map<String,Integer> columnIndexMap = new Map<String,Integer>();

    /**
     * @description Method to populate the column index map with the starting index and number of columns
     * @param  startIndex      The starting index of the column
     * @param  numberOfColumns The number of columns to be added to the map
     */
    public static void populateColumnIndexMap(Integer startIndex, Integer numberOfColumns){
        for(Integer i=startIndex;i<numberOfColumns;i++){
            columnIndexMap.put(CommonUtil.columnNameFromColumnIndex(i),i);
        }
    }


    /**
     * @description Method to get the letter part of an input cellname i.e. ALL100 returns ALL
     * @param  cellName The cell name to extract the column letter from
     * @return          The column letter in the form of a string
     */
    public static String columnNameFromCellName(String cellName){
        Matcher m = Pattern.compile(Constant.CELL_NAME_REGEX).matcher(cellName);
        if(!m.find()){
            throw new ParseException(String.format(Constant.MSG_INVALID_CELL_NAME, new String[]{cellName}));
        }
        return m.group(1);
    }
}
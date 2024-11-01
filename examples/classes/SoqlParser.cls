/**
 * @author         Justus van den Berg (jfwberg@gmail.com)
 * @date           August 2024
 * @copyright      (c) 2024 Justus van den Berg
 * @license        MIT (See LICENSE file in the project root)
 * @description    Class that parses SOQL Result JSON with nested child records and splits the results
 *                 according to sObjectType. Each sObjectType in the result gets it's own list of records
 *                 with corresponding metadata (field and relationship info)
 * 
 * @false-positive OneDeclarationPerLine    Only occurs in for loops for optimization purposes
 * @false-positive ExcessiveParameterList   I need to find a nicer way for the massive amount of parameters.
 *
 * @version 2.0    Version 2.0 adds metadata related to the sObject fields and relationships
 *                 By adding the metadata we can create output that can be ingested by external
 *                 solutions to create a point in time data model with the basic sObject Info
 *
 * @use case       The main use case is to transform SOQL query results that can be used
 *                 for exporting CSV or XLSX data that require individual tables for child lists.
 *                 A secondary use is to store related data for archiving purposes where a
 *                 potential reverse loading of related data might be required using an sObject Tree.
 */
@SuppressWarnings('PMD.OneDeclarationPerLine, PMD.ExcessiveParameterList')
public with sharing class SoqlParser {

    /** **************************************************************************************************** **
     **                                          PRIVATE CONSTANTS                                           **
     ** **************************************************************************************************** **/
    // Parameter for the relationship path, constant to prevent repetition of specifying a variable
    private final static String RELATIONSHIP_ATTRIBUTE = 'Relationship';

    // Parameter names to be ignored when creating the data table
    private final static Set<String> ATTRIBUTE_FILTER  = new Set<String>{'attributes', 'url', 'done', 'totalSize', 'nextRecordsUrl'};

    
    /** **************************************************************************************************** **
     **                                          PRIVATE VARIABLES                                           **
     ** **************************************************************************************************** **/
    // Map to keep track of the index of the sObject index in the output list
    private transient static Map<String,Integer> sObjectTypeIndexMap = new Map<String,Integer>();

    // List that holds the schema sObject describes in memory
    private transient static Map<String, Schema.DescribeSObjectResult> sObjectDescribeMap = new Map<String,Schema.DescribeSObjectResult>();


    /** **************************************************************************************************** **
     **                                       PUBLIC SUPPORT METHODS                                         **
     ** **************************************************************************************************** **/
    /**
     * @description Create table from an untyped Object list
     * @param input A list of Objects to traverse
     * @return      A list of flattened JSON records with metadata details (SObjectType classes)
     * @note        This is the main method to be called when invoking this class
     */
    public static SObjectType[] create(Object[] input){

        // Create the result object that will store the traversal results
        SObjectType[] output = new SObjectType[]{};

        // Input validation, if the input is null, return an empty list
        if(!(input?.isEmpty() == false)){return output;}

        // Get the SObject type for the parent table by extracting the attribute from the first record
        String parentSObjectName = String.valueOf((
            (Map<String,Object>)(((Map<String,Object>) input[0])
            .get('attributes')))
            .get('type')
        );

        // Populate the data by traversing the input and add the data to the output
        // We set the sObject name as the starting path and sObject
        traverseList(
            input,
            output,
            parentSObjectName,      // sObject Type
            new Node(
                null,               // Relationship path
                parentSObjectName,  // Current sObjectType
                null,               // Current field
                0                   // relationshipDepth
            )
        );

        // Return the flat table
        return output;
    }


    /** **************************************************************************************************** **
     **                                       PRIVATE TRAVERSE METHODS                                       **
     ** **************************************************************************************************** **/
    /**
     * @description Traverse an input list and add the result to the output
     * @param input       The input to traverse, this is the first record list
     * @param output      The Result object list containing all the data and metadata
     * @param sObjectType Starting path of a new sObject level traversal (sObject Name)
     * @param node        Current node in the traversal
     */
    private static void traverseList(Object[] input, SObjectType[] output, String sObjectType, Node node){

        // Get the list index for the list sObjectType
        Integer index = getSObjectTypeIndex(output, sObjectType);

        // Add the current relationship to the list of relationships for the sObjectType
        populateRelationshipMetadata(output, index, node);

        // Having this type casting again seems redundant, but doing the check in here
        // saves having to do it twice, what on 10k+ statements gives a small performance improvement
        for(Integer i=0, max=input.size();i<max;i++){

            // Create a new record to combine the values in the list
            Map<String,Object> record = new Map<String,Object>();

            // Get the object for improved performance
            Object o = input[i];

            // We only need to handle a child object if it's an object map
            if(o instanceof Map<String,Object>){
                traverseMap(
                    (Map<String,Object>) o,
                    output,
                    record,
                    sObjectType,
                    node
                );
            }

            // After the traversal is complete, Add the full record to the table
            // Make sure it is not empty to prevent empty rows
            if(!record.isEmpty()){

                // Add the relationship path to the attributes so we can track the origin of the record
                if(node.relationshipPath != null){
                    record.put(RELATIONSHIP_ATTRIBUTE, node.relationshipPath);
                }

                // Populate the field metadata for the record
                // We have to check each record as when a value is null in a query result
                // that value is not part of the payload, so we need to check all the fields
                populateFieldMetadata(output, record, index, sObjectType);

                // Add the full record to the matching list at the index of the result object
                output[index].records.add(record);
            }
        }
    }


    /**
     * @description Method to traverse a map and add the results to a result output
     * @param input     The input to traverse
     * @param output    The Result object list containing all the data and metadata
     * @param record    The current record in the traverse path
     * @param fieldPath The location in the traverse path for a field (i.e. Owner.Profile.Name)
     * @param node      Current node in the traversal
     */
    @SuppressWarnings('PMD.CognitiveComplexity')
    private static void traverseMap(Map<String,Object> input, SObjectType[] output, Map<String,Object> record, String fieldPath, Node node){

        // Iterate all the values in the input
        for(String key : input.keySet() ){

            // Continue if an attribute needs to be ignored
            if(ATTRIBUTE_FILTER.contains(key)){continue;}

            // Get the object for improved performance
            Object o = input.get(key);

            // If the child node is an object, (i.e. owner.name), traverse as map
            if(o instanceof Map<String,Object>){

                // Traverse the map, update the field path to include the child path
                traverseMap(
                    (Map<String,Object>) o,
                    output,
                    record,
                    String.isBlank(fieldPath) ? key : fieldPath + '.' + key,
                    new Node(
                        node.relationshipPath,
                        node.currentSObject,
                        node.currentRelationshipField,
                        node.relationshipDepth
                    )
                );

            // If the child node is an object list, traverse as list
            }else if(o instanceof Object[]){

                // Update the node with relationship info
                String relationshipPath = setNodeRelationshipInfo(fieldPath, node);

                // Traverse the list, update the relationship path and the relationship depth
                traverseList(
                    (Object[]) o,
                    output,
                    node.currentSObject,
                    new Node(
                        relationshipPath,
                        node.currentSObject,
                        node.currentRelationshipField,
                        node.relationshipDepth + 1
                    )
                );

            // If it's not a map or a list, it must a value, so add the value to record
            }else{
                record.put(
                    (String.isBlank(fieldPath) ? key : fieldPath + '.' + key ).removeStart(node.currentSObject + '.'),
                    o
                );
            }
        }
    }


    /** **************************************************************************************************** **
     **                                       PRIVATE SUPPORT METHODS                                       **
     ** **************************************************************************************************** **/
    /**
     * @description Method to get the sObject Index for a specific sObjectType
     *              If the type does not exist yet, this method will create a new
     *              SObjectType instance and add it to the output list. It also 
     *              adds the sObject describe information to the sObject Describe Map
     * @param  output  The Result object list containing all the data and metadata
     * @param  sObjectType   Starting path of a new sObject level traversal (sObject Name)
     * @return The index of the sObjectType in the output list
     */
    private static Integer getSObjectTypeIndex(SObjectType[] output, String sObjectType){
        
        // If we have a new sObject, create a SObjectType for that sObject
        // Add the index to the list of sObjects, smaller objects makes handling the results faster
        if(!sObjectTypeIndexMap.containsKey(sObjectType)){

            // Add the index based on the list size (zero index)
            sObjectTypeIndexMap.put(sObjectType, output.size());
            
            // Add a new result sObject, if it's value 0, it's the parent sObject
            output.add(new SObjectType(
                sObjectType,
                output.size() == 0
            ));

            // Describe the sObject and add the describe to the describe map
            sObjectDescribeMap.put(
                sObjectType,
                ((SObject) Type.forName(sObjectType).newInstance())
                .getSObjectType()
                .getDescribe()
            );
        }

        // Return the list index
        return sObjectTypeIndexMap.get(sObjectType);
    }


    /**
     * @description  Method that adds a relationship object to a ResultsObject
     *               class instance at the specified index
     * @param output The Result object list containing all the data and metadata
     * @param index  The index of the current sObjectType in the output
     * @param node   Current node in the traversal
     */
    private static void populateRelationshipMetadata(SObjectType[] output, Integer index, Node node){
        
        // If the child relationships do not contain the relationship path, add the relationship
        if(node.relationshipPath != null && !output[index].relationships.containsKey(node.relationshipPath)){
            output[index].relationships.put(
                node.relationshipPath,
                new Relationship(
                    node.currentRelationshipField,
                    node.relationshipDepth
                )
            );
        }
    }


    /**
     * @description Method that adds a relationship object to a ResultsObject
     *              class instance at the specified index
     * @param output      The Result object list containing all the data and metadata
     * @param record      The current record in the traverse path
     * @param index       The index of the current sObjectType in the output
     * @param sObjectType Starting path of a new sObject level traversal (sObject Name)
     */
    private static void populateFieldMetadata(SObjectType[] output, Map<String,Object> record, Integer index, String sObjectType){
        
        // Extract the fieldmap for performance improvement
        Map<String,Schema.SObjectField> fieldMap = sObjectDescribeMap?.get(sObjectType).fields.getMap();

        // Check for each field if we need to add it to the list
        for(String field : record.keySet()){
            if(!output[index].fields.containsKey(field)){

                output[index].fields.put(
                    field,
                    new ResultSObjectField(
                        field,
                        fieldMap?.get(field.toLowerCase())?.getDescribe()
                    )
                );
            }
        }
    }


    /**
     * @description Method that populates the current node with relationship info
     * @param  fieldPath Full path of the current node in the traversal (i.e. Owner.Profile.Name)
     * @param  node      Current node in the traversal
     * @return      The relationship path for the current node
     */
    private static String setNodeRelationshipInfo(String fieldPath, Node node){

        // Extract method for performance so it's not looped
        String relationshipName = fieldPath.substringAfterLast('.');

        // Iterate all the child relationships until the rihgt one is found
        for(Schema.ChildRelationship cr : sObjectDescribeMap.get(node.currentSObject).childrelationships  ){
            if(cr.relationshipname == relationshipName){
                node.currentSObject = String.valueOf(cr.childsobject);
                node.currentRelationshipField = String.valueOf(cr.field);
                break;
            }
        }

        // Keep track of the relationship name based on the last item in the field path
        return  (node.relationshipPath ?? '') +
                (String.isNotBlank(node.relationshipPath) ? '.' :  '') +
                (fieldPath.contains('.') ? (fieldPath?.substringAfterLast('.') ?? '') : fieldPath);
    }


    /** **************************************************************************************************** **
     **                                            PUBLIC CLASSES                                           **
     ** **************************************************************************************************** **/
    /**
     * @description Class representing the results for a single sObjectType
     *              It contains both the data and metadata
     */
    public class SObjectType{

        // Indicator if this is the parent sObject
        public Boolean isParent;

        // sObjectName / Type
        public String sObjectType;

        // Field info
        public Map<String, ResultSObjectField> fields = new Map<String,ResultSObjectField>();

        // Map with relationship objects
        public Map<String,Relationship> relationships = new Map<String,Relationship>();

        // List of records
        public List<Map<String,Object>> records       = new List<Map<String,Object>>();
       
        /**
         * @description Default constructor
         * @param sObjectType The sObject Type of this sObject
         * @param isParent    Indicator if this is the parent object
         */
        public SObjectType(String sObjectType, Boolean isParent){
            this.sObjectType = sObjectType;
            this.isParent    = isParent;
        }
    }


    /**
     * @description Class representing a relationship
     */
    public class Relationship{

        // Public variables
        public String  field;
        public Integer depth;

        /**
         * @description Default constructor
         * @param field The name of the sObject relationship field
         * @param depth The relationship depth
         */
        public Relationship(String field, Integer depth){
            this.field = field;
            this.depth = depth;
        }
    }


    /**
     * @description Class representing field info
     */
    public class ResultSObjectField{

        // Public variables
        public String  label;
        public String  type;
        public Boolean externalId;
        public String  relationshipName;
        public Boolean masterDetail;
        public Boolean readOnly;

        /**
         * @description Default constructor
         * @param name The name of the sObject field
         * @param dfr  The Schema.DescribeFieldResult for the field in question
         */
        public ResultSObjectField(String name, Schema.DescribeFieldResult dfr){
            this.label            = dfr?.label                      ?? name;
            this.type             = dfr?.type.name()                ?? 'UNKNOWN';
            this.relationshipName = (dfr?.relationshipname  != null) ? dfr?.relationshipname : null;
            this.masterDetail     = (dfr?.relationshiporder != null) ? true : null;
            this.readOnly         = !(dfr?.createable               ?? false);
            this.externalId       = dfr?.externalid                 ?? false;
        }
    }


    /** **************************************************************************************************** **
     **                                            PRIVATE CLASSES                                           **
     ** **************************************************************************************************** **/
    /**
     * @description A class that holds data for the current node whilst traversing a tree
     *              This is used to prevent having a lot of method variables being passed
     *              back and forth between the traversal methods.
     */
    private class Node{

        // Public variables
        public String  relationshipPath;
        public String  currentSObject;
        public String  currentRelationshipField;
        public Integer relationshipDepth;

         /**
         * @description Default constructor
         * @param relationshipPath         The full relationship path
         * @param currentSObject           The current object in the traversal
         * @param currentRelationshipField The current relationship field in the traversal
         * @param relationshipDepth        The current relationship depth in the traversal
         */
        public Node(String relationshipPath, String currentSObject, String currentRelationshipField, Integer relationshipDepth){
            this.relationshipPath = relationshipPath;
            this.currentSObject = currentSObject;
            this.currentRelationshipField = currentRelationshipField;
            this.relationshipDepth = relationshipDepth;
        }
    }
}
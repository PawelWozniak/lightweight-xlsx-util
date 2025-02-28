/**
 * @author      Justus van den Berg (jfwberg@gmail.com)
 * @date        May 2022
 * @copyright   (c) 2024 Justus van den Berg
 * @license     MIT (See LICENSE file in the project root)
 * @description Execute anonymous script to test a dependency topological
 *              sort against all sObjects that have data in the org and their
 *              lookup fields. Sorted based on the required loading order.
 *              Please note this is just a quickly put together example script
 *              that can be optimized and customized.
 *
 * @use case    The main use case is to sort sObjects and their relationships to create
 *              the order of loading during data migrations or data restores
 * 
 * @relatedCode https://www.thiscodeworks.com/salesforce-apex-topological-sort-automatically-define-sobject-data-migration-loading-order/66fdbd5592e2590014ee4528
 * 
 * @blog        https://medium.com/@justusvandenberg/programmatically-find-the-order-to-load-salesforce-objects-in-a-data-migration-using-apex-1f65841531fb
 */
@SuppressWarnings('PMD.CognitiveComplexity, PMD.ApexSuggestUsingNamedCred')
public with sharing class SortSObjectsWithData{

    // Keep track of all objects that have been described
    private static transient Set<Schema.SObjectType> finishedSObjectTypes = new Set<Schema.SObjectType>{};
    
    // Map to auto populate missing sObjects
    private static transient Set<Schema.SObjectType> missingSObjectTypes = new Set<Schema.SObjectType>{};
    
    // The node dependency map
    private static transient Map<Schema.SObjectType, Set<Schema.SObjectType>> nodeDependenciesMap = new Map<Schema.SObjectType,Set<Schema.SObjectType>>();
    
   
    /**
     * @description Method to get all the sorted nodes with info
     */
    public static Set<SortUtil.Node> getSortedNodes(){
    
        // Execute the main logic
        populateSObjectNodes(getSObjectsThatHaveData());
        
        // Populate the dependencies
        populateDependencies();

        // Execute the topological search against all sObjects that have data in them
        return SortUtil.topologicalSort(convertSchemaMapToObjectMap());
    }
    
 
    /**
     * @description Method that calls the limits API and gets all objects and their record count
     *              Add the records with a count > 0 to the output list
     * @return      A list of all SObjects that contain at least 1 record
     */
    private static String[] getSObjectsThatHaveData(){
    
        // List of the output sObjects
        String[] sObjectsWithData = new String[]{};
    
        // Create new request
        HttpRequest request = new HttpRequest();
        request.setEndpoint(URL.getOrgDomainUrl().toExternalForm() + '/services/data/v63.0/limits/recordCount');
        request.setMethod('GET');
        request.setHeader('Content-Type', 'application/json;charset=UTF-8');
        request.setHeader('Authorization', 'Bearer ' + userInfo.getSessionId());
        HttpResponse response = new http().send(request);
    
        // Check response code
        if (response.getStatusCode() != 200) {
            throw new StringException(response.getBody());
        }
    
        // Populate storage list, note fastest JSON parsing happens when directly going into a constructor instead of creating a variable assignment first
        for(Object obj : (Object[]) ((Map<String, Object>) JSON.deserializeUntyped(response.getBody())).get('sObjects')){
            populateSObject((Map<String,Object>) obj, sObjectsWithData);
        }
    
        // return the list with data
        return sObjectsWithData;
    }


    /**
     * @description Method to add an object name to the list if there is more than one record
     * @param objMap           The object map with the limit details
     * @param sObjectsWithData String list of objects with data
     */
    private static void populateSObject(Map<String,Object> objMap,  String[] sObjectsWithData){
        if( (Integer) objMap.get('count') > 0){
            sObjectsWithData.add((String) objMap.get('name'));
        }
    }
    
    
    /**
     * @description Method to populate the sObject nodes
     * @param sObjectName A list of sObject API Names
     */
    private static void populateSObjectNodes(String[] sObjectNames){
        // Create the base map with empty sObject dependency sets
        for(String sObjectName : sObjectNames){
            
            // Describe the sObject
            Schema.DescribeSObjectResult dsor = ( (SObject) Type.forName('Schema.' + sObjectName).newInstance()).getSObjectType().getDescribe();
            
            // We only want writable and updateable sObjects that have a keyprefix
            // otherwise we cannot load the data, so skip otherwise
            if(dsor.createable == false || dsor.updateable == false || dsor.keyprefix == null){
                continue;
            }
    
            // Populate the node dependencies map
            nodeDependenciesMap.put(
                dsor.sobjecttype,
                new Set<Schema.SObjectType>{}
            );
        }
    }
    
    
    /**
     * @description Method to populate the sObject dependencies recursively
     */
    private static void populateDependencies(){
        
        // Iterate all the object types in the dependencies map
        for(Schema.SObjectType sot : nodeDependenciesMap.keySet()){
    
            // Only describe sObjects once
            if(finishedSObjectTypes.contains(sot)){
                continue;
            }
    
            for(Schema.SObjectField sof :  sot.getDescribe().fields.getMap().values()){
                
                // Describe the field
                Schema.DescribeFieldResult dfr = sof.getDescribe();
                
                // Skip everything except lookup fields
                if(dfr.type != Schema.DisplayType.REFERENCE){
                    continue;
                }
    
                // Get the object describe
                Schema.DescribeSObjectResult dsor = dfr.referenceto[0].getDescribe();
    
                // Also for the related describes we only want writable objects
                // This filters out things like history objects and RecordTypes
                if(dsor.createable == false || dsor.updateable == false || dsor.keyprefix == null){
                    continue;
                }
    
                // Add the lookup field as a dependency
                nodeDependenciesMap.get(sot).add(dsor.sobjecttype);
                
                // Dependency is missing, so add it that we can add it later
                if(!nodeDependenciesMap.containsKey(dsor.sobjecttype)){
                    nodeDependenciesMap.put(dsor.sobjecttype, new Set<Schema.SObjectType>{});
                    missingSObjectTypes.add(dsor.sobjecttype);
                }
    
                // If the the sot has been added so can be removed from the missing SOTs
                if(missingSObjectTypes.contains(sot)){
                    missingSObjectTypes.remove(sot);
                }
            }
    
            // Flat that sObject describe has finished
            finishedSObjectTypes.add(sot);
        }
    
        // Recursively call self
        if(!missingSObjectTypes.isEmpty()){
            populateDependencies();
        }
    }
    
    
    /**
     * @description Method to convert all Schema.SObjectTypes to Strings
     *              Fixes the "bug" where the self reference should not 
     *              be the first dependency
     * @return      A node/dependency map (object,Set<Object>)
     */
    private static Map<Object,Set<Object>> convertSchemaMapToObjectMap(){
        
        // Output map
        Map<Object,Set<Object>> objectMap = new Map<Object,Set<Object>>();
        
        // Add the string values for each sObject
        for(Schema.SObjectType sot : nodeDependenciesMap.keySet()){
            
            // Set for holding the dependencies
            Set<Object> dependencies = new Set<Object>{};
            
            // Counter
            Integer i = 0;
    
            // Indicate if relationship to self needs to be added to the end
            // This fixes a bug where if the first dependency is to itself all other get ignored.
            Boolean addSelfToEnd = false;
    
            // Convert dependencies
            for(Schema.SObjectType dependency : nodeDependenciesMap.get(sot)){
                if(i==0 && (dependency == sot)){
                    addSelfToEnd = true;
                    i++;
                    continue;
                }
                
                dependencies.add(String.valueOf(dependency));
                i++;
            }
    
            // Add self to end of dependencies
            if(addSelfToEnd){
                dependencies.add(String.valueOf(sot));
                addSelfToEnd = false;
            }
    
            // Add the object type and dependencies
            objectMap.put(String.valueOf(sot),dependencies);
        }
    
        // Return the converted map
        return objectMap;
    }
}
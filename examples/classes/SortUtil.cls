/**
 * @author      Justus van den Berg (jfwberg@gmail.com)
 * @date        May 2022
 * @copyright   (c) 2024 Justus van den Berg
 * @license     MIT (See LICENSE file in the project root)
 * @description Class contains a basic topological (graph) sort alogrithm for Apex
 *              It's not 100% perfect, but does the job for most simple graph sorts.
 *              It has two additional methods, one for creating an ordered list of all
 *              nodes in the graph sorted by dependency and a second one to create
 *              a report of any errors or issues with the graph
 *
 * @use case    The main use case is to sort sObjects and their relationships to create
 *              the order of loading during data migrations or data restores
 * 
 * @relatedCode https://www.thiscodeworks.com/salesforce-apex-get-all-sobjects-with-data-and-sort-them-topologically/66fdbe1192e2590014ee4673
 * @blog        https://medium.com/@justusvandenberg/programmatically-find-the-order-to-load-salesforce-objects-in-a-data-migration-using-apex-1f65841531fb
 */
@SuppressWarnings('PMD.ExcessiveParameterList')
public with sharing class SortUtil {
 
    // Warning message templates
    private static final String SELF_DEPENDENCY_WARNING	   = 'Node "{0}" has a dependency on itself';
    private static final String DUO_DEPENDENCY_WARNING	   = 'Node "{0}" and "{1}" have a dependency on each other.';
    private static final String CYL_DEPENDENCY_WARNING	   = 'Node "{0}" has a cylindrical relationship on path "{1}".';
    private static final String MISSING_DEPENDENCY_WARNING = 'Node "{0}" depends on node "{1}" but Node "{1}" does not exist in the dependencies Map';
 
    // Create a map to hold any warnings
    private static Map<Object, String[]> warningMap = new Map<Object, String[]>();
 
 
    /**
     * @description Sort a graph topologically
     * @param nodeDependenciesMap the map containing the node and all of it's dependencies
     * @return      A sorted list of the nodes based on the dependency
     */
    public static Set<Node> topologicalSort(Map<Object,Set<Object>> nodeDependenciesMap){
        
        // Validate graph first to create the initial warning messages
        // Execute the topological sorting and return the details
        return new TopologicalSort(
            nodeDependenciesMap,
            validateGraph(nodeDependenciesMap)
        ).getSortedResult();
    }
 
 
    /**
     * @description A methods that analyses the data map and generates
     *              warnings for missing, circular or self dependencies
     * @param nodeDependenciesMap The main node/dependencies map
     * @return      A map with (potential) warnings for each node
     */
    private static Map<Object, String[]> validateGraph(Map<Object,Set<Object>> nodeDependenciesMap){
 
        // Iterate each node
        for(Object node : nodeDependenciesMap.keySet()){
 
            for(Object dependency : nodeDependenciesMap.get(node)){
 
                // If a referenced dependency does not exist in the data map add a warning
                if(!nodeDependenciesMap.containsKey(dependency)){
                    addWarning(
                        node, 
                        String.format(MISSING_DEPENDENCY_WARNING, new String[]{
                            String.valueOf(node),
                            String.valueOf(dependency)
                        })
                    );
                    
                    // If the dependend node does not exists in the data map there is no point to check for self or circular dependencies
                    // so continue to the next dependency
                    continue;
                }
 
                // If there is a dependency to the same node, add a warning
                if(node == dependency){
                    addWarning(
                        node,
                        String.format(
                            SELF_DEPENDENCY_WARNING,
                            new String[]{String.valueOf(node)}
                        )
                    );
 
                    // If the dependend node is self referencing, there is no need to look for a duo dependency as
                    // technically it already is a duo reference
                    continue;
                }
                    
                // If there is a duo dependency between the dependency and the node, add a warning message (ie they point at each other)
                if(nodeDependenciesMap.get(dependency).contains(node)){
                    addWarning(
                        node,
                        String.format(
                            DUO_DEPENDENCY_WARNING,
                            new String[]{String.valueOf(node),String.valueOf(dependency)}
                        )
                    );
                }
            }
        }
        return warningMap;
    }
 
 
    /**
     * @description Method that adds a warning to a specific node in the warning maps
     * @param node       The node object
     * @param warning    The warning message
     */
    private static void addWarning(Object node, String warning){
        
        // Create new map entry if it doesnt exist
        if(!warningMap.containsKey(node)){
           warningMap.put(node,new String[]{});
        }
 
        // Add warning to the map
        warningMap.get(node).add(warning);
    }
 
 
    /**
     * description Class to perform a topological sort
     */
    private with sharing class TopologicalSort{
        
        // Start the order for the nodes. We use a zro index so start at -1
        private Integer nodeOrder = -1;
 
        // A set with the sorted result
        private transient Set<Node> sortedResult = new Set<Node>{};
 
        // The map containing the warnings
        private transient Map<Object, String[]> warningsMap = new Map<Object, String[]>();
 
 
        /**
         * @description Constructor that that the map with nodes and their dependencies
         * @param nodeDependenciesMap The main input node / depencencies map
         * @param warningsMap         The map containing the warning messages
         */
        private TopologicalSort(Map<Object, Set<Object>> nodeDependenciesMap, Map<Object, String[]> warningsMap){
            
            // Assign the warning map
            this.warningsMap = warningsMap;
 
            // Start the sorting
            visit(
                nodeDependenciesMap,
                nodeDependenciesMap.keySet(),
                new Set<Object>{},
                new Set<Object>{},
                ''
            );
        }
 
 
        /**
         * @description Method to get the sorted results
         * @return The result from the topological sort
         */
        private Set<Node> getSortedResult(){
            return this.sortedResult;
        }
 
 
        /**
         * @description Method to pass through i.e. visit each node in the graph
         * @param nodeDependenciesMap The main map with the nodes and their dependencies
         * @param dependencies        The dependencies for this specific node
         * @param dead                Set with nodes that have been visited
         * @param pending             Set with nodes that have not been visited
         * @param nodePath            The path that has led to the node
         */
        @SuppressWarnings('PMD.CognitiveComplexity')
        private void visit(Map<Object,Set<Object>> nodeDependenciesMap, Set<Object> dependencies, Set<Object> dead, Set<Object> pending, String nodePath ){
            
            // If there are no dependencies nothing needs to happen
            if(dependencies == null){return;}
            
            // Check all dependencies for this node
            for(Object node : dependencies){
 
                String childNodePath = nodePath + (!String.isBlank(nodePath) ? + '.' : '') + node;
                
                // If the node has not been visited before
                if(!dead.contains(node)){
                    if (!pending.contains(node)){
                        pending.add(node);
                    }else{
                        // Circular relationship
                        addWarning(
                            node,
                            String.format(
                                CYL_DEPENDENCY_WARNING,
                                new String[]{String.valueOf(node),childNodePath}
                            )
                        );
                        return;
                    }
 
                    // recursively call this function for every child of the current node
                    visit(nodeDependenciesMap, nodeDependenciesMap.get(node), dead, pending, childNodePath);
 
                    // If the node is still pending, remove it
                    if (pending.contains(node)){
                        pending.remove(node);
                    }
 
                    // Indicade that this node has been visited
                    dead.add(node);
 
                    // Increment the order, it starts at -1, so the first node will be 0
                    this.nodeOrder++;
 
                    // Add the node to the end of the sorted list
                    this.sortedResult.add(new Node(
                        this.nodeOrder,
                        node,
                        childNodePath.split('\\.').size() -1,
                        childNodePath,
                        this.warningsMap.get(node)
                    ));   
                }
            }
        }
    }
 
 
    /**
     * @description Class that represents a node in a graph that can be topologically sorted
     */
    @SuppressWarnings('PMD.ExcessiveParameterList')
    public with sharing class Node implements comparable{
        
        // Public variables
        public Integer  order;
        public Object   node;
        public Integer  depth;
        public String   path;
        public String[] warnings;
 
        /**
         * @description The main constructor
         * @param order     The order in the graph
         * @param node      The node itself
         * @param depth     The node depth
         * @param path      The path it took to get to the node
         * @param warnings  List of warnings for the node
         */
        
        public Node(Integer order, Object node, Integer depth, String path, String[] warnings){
            this.order    = order;
            this.node     = node;
            this.depth    = depth;
            this.path     = path;
            this.warnings = warnings;
        }
 
        /**
         * @description Interface for compare method to sort based on order
         * @param compareTo The Node to compare against
         * @return      -1,0 or 1
         */
        public Integer compareTo (Object compareTo) {
            return (this.order > ((Node)compareTo).order) ? 1 : (this.order < ((Node)compareTo).order) ? -1 : 0;
        }
    }
    
  
}
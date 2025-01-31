/**
 * @author      Justus van den Berg (jfwberg@gmail.com)
 * @date        August 2024
 * @copyright   (c) 2024 Justus van den Berg
 * @license     MIT (See LICENSE file in the project root)
 * @description Test Data generator for SoqlTableParser & SoqlMultiTableParser
 * @note        Create the test data, ideally run 1 by 1 to prevent govenor limits
 *              Run accounts and contacts first. Cases and opportunities require both
 *
 * - https://medium.com/@justusvandenberg/dynamically-handle-salesforce-soql-subquery-response-data-using-apex-8130bd0622aa
 * - https://www.thiscodeworks.com/apex-soqltableparser-soql-subqueries-to-a-single-flat-table/66ba81bc2e62a20014b9150c
 * - https://www.thiscodeworks.com/apex-soqlmultitableparser-soql-subqueries-to-multiple-individual-tables/66ba816f2e62a20014b913aa
 */
insertAccounts(0,10);
insertContacts(0,10,0,10);
insertCases(0,10,0,10);
insertOpportunities(0,10,0,10);
insertOpportunitContactRoles(0,10,0,10);
 
/**
 * @description Method to create Basic Test Contacts
 */
public void insertAccounts(Integer accountOffset, Integer numberOfAccounts){
    
    // List to store the new cases
    Account[] accounts = new Account[]{};
 
    // Loop for to create all top level accounts
    for (Integer i = accountOffset; i < numberOfAccounts; i++) {
    
        // A postfix to keep track of what item we're dealing with
        String postfix = String.valueOf(i+1).leftPad(4,'0');
        
        // Create a new account
        accounts.add(new Account(
            Name = 'LWT - Test - Account - ' + postfix
        ));
    }
    insert accounts;
}
 
 
/**
 * @description Method to create Basic Test Contacts
 */
public void insertContacts(Integer accountOffset, Integer numberOfAccounts, Integer contactOffset,Integer numberOfContacts){
    
    // Get the account and contact data
    Account[] accounts = [SELECT Id FROM Account WHERE Name LIKE 'LWT - Test - %' ORDER BY Name ASC LIMIT :numberOfAccounts];
    
    // Basic error handling
    if(numberOfAccounts != accounts.size()){throw new StringException('Number of accounts does not match number returned by the query');}
    
    // List to store the new contacts
    Contact[] contacts = new Contact[]{};
 
    // Iterate the top level accounts
    for(Integer i=accountOffset; i<numberOfAccounts; i++){
        
        // Create a number of contacts for each account
        for(Integer j=contactOffset; j < numberOfContacts; j++){
        
            // Postfix to keep track of where we are
            String postfix = String.valueOf(i+1).leftPad(4,'0') + ' - ' + String.valueOf(j+1).leftPad(4,'0');
            
            // Add a new contact to the list
            contacts.add(new Contact(
                AccountId = accounts[i].Id,
                FirstName = 'LWT - Test - ' + postfix,
                LastName  = 'Contact - '    + postfix
            ));
        }
    }
    insert contacts;
}
 
 
/**
 * @description Method to create Basic Test Cases
 */
public void insertCases(Integer offset, Integer numberOfAccounts, Integer contactOffset, Integer numberOfContacts){
 
    // Get the account and contact data
    Account[] accounts = [SELECT Id FROM Account WHERE Name LIKE 'LWT - Test - %' ORDER BY Name ASC LIMIT :numberOfAccounts];
    Contact[] contacts = [SELECT Id FROM Contact WHERE Name LIKE 'LWT - Test - %' ORDER BY Name ASC LIMIT :numberOfContacts];
 
    // Basic error handling
    if(numberOfAccounts != accounts.size()){throw new StringException('Number of accounts does not match number returned by the query');}
    if(numberOfContacts != contacts.size()){throw new StringException('Number of contacts does not match number returned by the query');}
 
    // List to store the new cases
    Case[] cases = new Case[]{};
 
    // Iterate the top level accounts
    for(Integer i=offset; i<numberOfAccounts; i++){
        
        // Create a case for each contact in the account
        for(Integer j=offset; j<numberOfContacts; j++){
        
            // Postfix to keep track of where we are
            String postfix = String.valueOf(i+1).leftPad(4,'0') + ' - ' + String.valueOf(j+1).leftPad(4,'0');
            
            // Add a new case
            cases.add(new Case(
                AccountId = accounts[i].Id,
                ContactId = contacts[j].Id,
                Subject   = 'LWT - Test - Case - ' + postfix
            ));
        }
    }
    insert cases;
}
 
 
/**
 * @description Method to create Basic Test Opportunities
 */
public void insertOpportunities(Integer offset, Integer numberOfAccounts, Integer contactOffset, Integer numberOfContacts){
    
    // Get the account and contact data
    Account[] accounts = [SELECT Id FROM Account WHERE Name LIKE 'LWT - Test - %' ORDER BY Name ASC LIMIT :numberOfAccounts];
    Contact[] contacts = [SELECT Id FROM Contact WHERE Name LIKE 'LWT - Test - %' ORDER BY Name ASC LIMIT :numberOfContacts];
 
    // Basic error handling
    if(numberOfAccounts != accounts.size()){throw new StringException('Number of accounts does not match number returned by the query');}
    if(numberOfContacts != contacts.size()){throw new StringException('Number of contacts does not match number returned by the query');}
    
    // List to store the new cases
    Opportunity[] opportunities = new Opportunity[]{};
 
    // Iterate the top level accounts
    for(Integer i=offset; i<numberOfAccounts; i++){
        
        // Create an opportunity for each contact in the account
        for(Integer j=offset; j<numberOfContacts; j++){
        
            // Postfix to keep track of where we are
            String postfix = String.valueOf(i+1).leftPad(4,'0') + ' - ' + String.valueOf(j+1).leftPad(4,'0');
            
            // Add a new opportunity to the list
            opportunities.add(new Opportunity(
                AccountId      = accounts[i].Id,
                ContactId      = contacts[j].Id,
                Name           = 'LWT - Test - Opportunity - ' + postfix,
                StageName      = (Math.mod(j,2) == 0) ? 'New' : 'Closed/Won',
                CloseDate      = Date.today().addDays(j)
            ));
        }
    }
    insert opportunities;
}
 
 
/**
 * @description Method to create Basic Test Opportunities
 */
public void insertOpportunitContactRoles(Integer offset, Integer numberOfAccounts, Integer contactOffset, Integer numberOfContacts){
    
   // Get the account and contact data
    Account[]     accounts      = [SELECT Id FROM Account     WHERE Name LIKE 'LWT - Test - %' ORDER BY Name ASC LIMIT :numberOfAccounts];
    Contact[]     contacts      = [SELECT Id FROM Contact     WHERE Name LIKE 'LWT - Test - %' ORDER BY Name ASC LIMIT :numberOfContacts];
    Opportunity[] opportunities = [SELECT Id FROM Opportunity WHERE Name LIKE 'LWT - Test - %' ORDER BY Name ASC LIMIT :numberOfContacts];
 
    // Basic error handling
    if(numberOfAccounts != accounts.size()){     throw new StringException('Number of accounts does not match number returned by the query');}
    if(numberOfContacts != contacts.size()){     throw new StringException('Number of contacts does not match number returned by the query');}
    if(numberOfContacts != opportunities.size()){throw new StringException('Number of opportunities does not match number returned by the query');}
    
    // List to store the new cases
    OpportunityContactRole[] ocRoles = new OpportunityContactRole[]{};
 
    // Iterate the top level accounts
    for(Integer i=offset; i<numberOfAccounts; i++){
        
        // Create an opportunity for each contact in the account
        for(Integer j=contactOffset; j<numberOfContacts; j++){
        
            // Add a new opportunity to the list
            ocRoles.add(new OpportunityContactRole(
                ContactId      = contacts[j].Id,
                OpportunityId  = opportunities[j].Id
            ));
        }
    }
    insert ocRoles;
}
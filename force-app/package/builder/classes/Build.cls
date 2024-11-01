/**
 * @author         Justus van den Berg (jfwberg@gmail.com)
 * @date           November 2024
 * @copyright      (c) 2023 Justus van den Berg
 * @license        MIT (See LICENSE file in the project root)
 * @description    Class that holds methods to create the XLSX file from a builder class instance
 * @false-positive AvoidGlobalModifier      This is a utility that is designed to be called from a
 *                                          managed package. It can be used for a user's own custom
 *                                          Implementation, so global is the way to open up this
 *                                          utility for global use.
 */
@SuppressWarnings('PMD.AvoidGlobalModifier')
global with sharing class Build {

    /**
     * @description Method to build the XLSX builder as a Document
     * @param b The builder class instance
     * @return      The XLSX as a Document object
     */
    global static Document asDocument(Builder b){

       // Sanatize the developer name and add X in case the number begins with a number
       String developerName = 'x_' + BuildUtil.getSanatizedFileName(b.fileName)
            ?.removeEnd(Constant.FULL_FILE_EXTENSION)
            ?.toLowerCase()
            ?.replaceAll(' ','_')
        ;

        // Create a new document
        Document d = new Document();
        d.Type              = Constant.FULL_FILE_EXTENSION.removeStart('.');
        d.Name              = b.fileName;
        d.DeveloperName     = developerName;
        d.Description       = b.coreProperties.description;
        d.FolderId          = Constant.DEF_FOLDER_ID;
        d.IsInternalUseOnly = true;
        d.IsPublic          = false;

        // Create the file body using the package builder
        d.body = XmlPackageBuilder.getPackage(b);

        // Combine the keywords and cut of after the keywords get too much
        d.keywords = (String.join(b.coreProperties.keywords,', ').length() > 255) ?
            String.join(b.coreProperties.keywords,', ').substring(0,255).substringBeforeLast(',') :
            String.join(b.coreProperties.keywords,', ')
        ;

        // Return the fully aconfigured document
        // You can update the folder if you want later on
        return d;
    }


    /**
     * @description Method to build the XLSX builder as a ContentVersion
     * @param b The builder class instance
     * @return      The XLSX as a ContentVersion object
     */
    global static ContentVersion asContentVersion(Builder b){

        ContentVersion cv = new ContentVersion();

        // Create a new content version
        cv.Title        = b.coreProperties.title;
        cv.description  = b.coreProperties.description;
        cv.PathOnClient = BuildUtil.getSanatizedFileName(b.fileName);

        // Create the file body using the package builder
        cv.VersionData  = XmlPackageBuilder.getPackage(b);
        return cv;
    }
}
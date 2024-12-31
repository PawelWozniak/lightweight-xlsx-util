

// As data map
Object[][][] xlsxDataArray = xlsx.Parse.toArray(
    new Compression.ZipReader(
        [SELECT body FROM Document WHERE Id = '015Qz000004jf7yIAA' LIMIT 1]?.Body
    ).getEntriesMap()
);

System.debug(JSON.serializePretty(xlsxDataArray));
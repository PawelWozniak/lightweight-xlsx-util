XmlStreamWriter x = new XmlStreamWriter();


xlsx.XmlElement.SheetPr sheetPr = new xlsx.XmlElement.SheetPr();


xlsx.XmlElement.Worksheet worksheet = new xlsx.XmlElement.Worksheet();
worksheet.sheetPr = sheetPr;


worksheet.cols = new xlsx.XmlElement.Column[]{
    new xlsx.XmlElement.Column(12.34),
    new xlsx.XmlElement.Column(56.78),
    new xlsx.XmlElement.Column(100.10),
    new xlsx.XmlElement.Column()
}; 


worksheet.writeXml(x);

System.debug(x.getXmlString());



/*

xlsx.XmlElement.Cell cell1 = new xlsx.XmlElement.Cell();
xlsx.XmlElement.Cell cell2 = new xlsx.XmlElement.Cell();
xlsx.XmlElement.Cell cell3 = new xlsx.XmlElement.Cell();
xlsx.XmlElement.Cell cell4 = new xlsx.XmlElement.Cell();


cell.v = new CellValue(1);




System.debug(x.getXmlString());


xlsx.Builder b = new xlsx.Builder();
b.addWorksheet('name');
b.getWorksheet('name').setName(.firstSheet);

// add new row at end
b.getWorksheet('name').addRow()

// add multiple rows
b.getWorksheet('name').addRow(5);

b.getWorksheet('name').getRow(index);

b.getWorksheet('name').addCell(rowIndex);

b.getWorksheet('name').getCell(rowIndex, columnIndex).styleIndex();

*/


# WordHelper
WordHelper is a simple java application for constructing MS Word document using Word Template input.
The code implements Apache POI. For more information about Apache POI: https://poi.apache.org/

Features
--------
1. Search and replace text in document template
2. Search and Update table that has matching column headers
 
Usage Guide
------------
//Create instance with template file path as input:  
WordHelper rp = new WordHelper("/path/to/template/file/");

//Update template field:  
rp.replaceText("##CompanyName##", "MYCORP");

//Update table:  
  //Create List variable to be used for searching table in template file:  
List<String> tbheads = Arrays.asList("First Column","Second Column","Third Column"); 

  //Sample Two dimensional List to fill the table rows:  
List<List<String>> tbdata = Arrays.asList(Arrays.asList("Row2Col1","Row2Col2","Row2Col3"),Arrays.asList("Row3Col1","Row3Col2","Row3Col3"),Arrays.asList("Row4Col1","Row4Col2","Row4Col3")); 

  //Call updateTable function to update Table in document template:  
rp.updateTable(tbheads,tbdata); 

//Save Document:  
rp.saveAs("/path/to/output/document.doc");

Word Template Creation Guide
----------------------------
- Use unique keyword in document template to be searched and replaced by application. Example: ##CompanyName##
- Prepare table with header columns. Empty data rows are not required, the application will append the rows as needed.

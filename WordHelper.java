package com.ibm.custom;
//maintained by Paulus Gunadi (paulus@sg.ibm.com)

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Arrays;
import java.util.Iterator;
import java.util.Calendar;
import java.text.DateFormat;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class WordHelper {
  private XWPFDocument document;

   public WordHelper(String strin) throws Exception {
     FileInputStream infile = new FileInputStream(new File(strin));
     document = new XWPFDocument(infile);
   }

   public XWPFDocument getDocument() {
     return document;
   }

   public void saveAs(String strout) throws Exception {
     FileOutputStream outfile = new FileOutputStream(new File(strout));
     document.write(outfile);
     outfile.close();
   }

   public void replaceText(String src, String dst) {
     int found = 0;

     List<XWPFParagraph> plist = document.getParagraphs();
     for (XWPFParagraph p : plist) {
       List<XWPFRun> rlist = p.getRuns();
       for (XWPFRun r : rlist) {
         if (r.text().equals(src)) {
           found = 1;
           r.setText(dst,0);
           return;
         }
       }
     }

     for (XWPFTable tbl : document.getTables()) {
       for (XWPFTableRow row : tbl.getRows()) {
         for (XWPFTableCell cell : row.getTableCells()) {
           for (XWPFParagraph p : cell.getParagraphs()) {
             for (XWPFRun r : p.getRuns()) {
               if (r.text().equals(src)) {
                 found = 1;
                 r.setText(dst,0);
                 return;
               }
             }
           }
         }
       }
     }
   }

   public void updateTable(List<String> tbheads, List<List <String>> tbdata) {
     int found = 0;

     for (XWPFTable tbl : document.getTables()) {
       List<XWPFTableRow> rows = tbl.getRows();
       if (!rows.isEmpty()) {
         //Validate Row Headers
         int isdiff = 0;
         XWPFTableRow headrow = rows.get(0);
         List<XWPFTableCell> cells = headrow.getTableCells();
         if (!cells.isEmpty()) {
           if (cells.size() == tbheads.size()) {
             for (int i=0; i < cells.size(); i++) {
               if (!cells.get(i).getText().equalsIgnoreCase(tbheads.get(i))) {
                 isdiff = 1;
                 break;
               }
             }
           } else {
             isdiff = 1;
           }
         } else {
           isdiff = 1;
         }
         if (isdiff == 0) {
           int rowsize = rows.size();
           int datasize = tbdata.size();
           int counter = 0;
           Iterator it = tbdata.iterator();
           while (it.hasNext()) {
             counter++;
             //Prepare Table
             XWPFTableRow row = null;
             if (rowsize > counter) {
               //UpdateRow
               row = rows.get(counter);
             } else {
               //AddRow
               row = tbl.createRow();
             }
             List<XWPFTableCell> datacells = row.getTableCells();
             int cellsize = datacells.size();
             //Get Data
             List <String> tbdatacells = (List <String>) it.next();
             int datacellsize = tbdatacells.size();

             //Assign Value
             for (int j=0; j < datacellsize; j++ ) {
               if (cellsize > j) {
                 //UpdateCell
                 row.getCell(j).setText(tbdatacells.get(j));
               } else {
                 //AddCell
                 row.addNewTableCell().setText(tbdatacells.get(j));
               }
             }
           }
         }
       }
     }
   }
   public static void main(String[] args)throws Exception {
     //Document Template
     WordHelper rp = new WordHelper(args[0]);
     XWPFDocument doc = rp.getDocument();

     DateFormat df = DateFormat.getInstance();

     rp.replaceText("##LetterNum##", "123/456/789");
     rp.replaceText("##CurrentDate##", df.format(Calendar.getInstance().getTime()));
     rp.replaceText("##CustomerName##", "IBM");
     rp.replaceText("##TicketID##", "ABC123");
     rp.replaceText("##ResolveDate##", df.format(Calendar.getInstance().getTime()));
     rp.replaceText("##SolutionDetails##", "Restart Server\n Recreate User\n");
     rp.replaceText("##Resolver##", "John Doe");

     //Set Table data
     List<String> tbheads = Arrays.asList("Circuit ID","Speed","Location");
     List<List<String>> tbdata = Arrays.asList(Arrays.asList("CI100","100 Mbps","Singapore"),Arrays.asList("CI210","1 Gbps","Kuala Lumpur"),Arrays.asList("CI320","10 Gbps","Bangkok"));
     rp.updateTable(tbheads,tbdata);
     //Write the Document in file system
     rp.saveAs(args[1]);
     System.out.println("document written successully");
    }
}

package com.sanitizer.DocxSanitize;

import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.opc.OPCPackage;

import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;

public class ExcelSanitize {
    public static void main(String[] args) {
        try {
            File file = new File("TestSanitize.xlsx");
            OPCPackage pkg = OPCPackage.open(file);

            XSSFWorkbook workbook = new XSSFWorkbook(pkg);

            int numberOfSheets = workbook.getNumberOfSheets();

            // Not Working as Apache POI doesn't support new version of Threaded comments in Excel!
            for(int i = 0 ; i < numberOfSheets; i++) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                Map<CellAddress, XSSFComment> comments = sheet.getCellComments();
                for(Map.Entry<CellAddress, XSSFComment> entry: comments.entrySet()) {
                    XSSFComment comment = entry.getValue();
                    comment.setAuthor(null);
                }
            }

            POIXMLProperties props = workbook.getProperties();
            POIXMLProperties.CoreProperties core = props.getCoreProperties();
            core.setCreator(null);               // Clears Author
            core.setTitle(null);                 // Clears Title
            core.setSubjectProperty(null);       // Clears Subject
            core.setDescription(null);           // Clears Comments/Description
            core.setKeywords(null);              // Clears Keywords
            core.setLastModifiedByUser(null);    // Clears 'Last saved by'
            core.setCategory(null);              // Clears Category

            POIXMLProperties.ExtendedProperties ext = props.getExtendedProperties();
            ext.setCompany(null);                // Clears Company
            ext.setManager(null);                // Clears Manager
            ext.setHyperlinkBase(null);          // Clears HyperLinkBase

            POIXMLProperties.CustomProperties custom = props.getCustomProperties();
            if (custom != null) {
                List<CTProperty> list = custom.getUnderlyingProperties().getPropertyList();
                list.clear();
            }

            FileOutputStream fos = new FileOutputStream("SanitizedSheet.xlsx");
            workbook.write(fos);
            fos.close();
            workbook.close();
            pkg.close();
        }
        catch(Exception e) {
            System.out.println(e.getMessage());
        }
    }
}

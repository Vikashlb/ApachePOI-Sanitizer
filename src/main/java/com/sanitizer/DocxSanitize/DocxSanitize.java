package com.sanitizer.DocxSanitize;

import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

public class DocxSanitize {
    public static void main(String[] args) {
        try {
            File file = new File("TestSanitize.docx");
            OPCPackage pkg = OPCPackage.open(file);
            XWPFDocument doc = new XWPFDocument(pkg);
            POIXMLProperties props = doc.getProperties();
            POIXMLProperties.CoreProperties core = props.getCoreProperties();
            core.setCreator("");               // Clears Author
            core.setTitle("");                 // Clears Title
            core.setSubjectProperty("");       // Clears Subject
            core.setDescription("");           // Clears Comments/Description
            core.setKeywords("");              // Clears Keywords
            core.setLastModifiedByUser("");    // Clears 'Last saved by'
            core.setCategory("");              // Clears Category
            core.setContentStatus("");         // Clears Status

            POIXMLProperties.ExtendedProperties ext = props.getExtendedProperties();
            ext.setCompany("");
            ext.setManager("");
            ext.setHyperlinkBase("");

            POIXMLProperties.CustomProperties custom = props.getCustomProperties();
            if(custom!= null) {
                List<CTProperty> list = custom.getUnderlyingProperties().getPropertyList();
                list.clear();
            }
            FileOutputStream fos = new FileOutputStream("Sanitized.docx");
            doc.write(fos);
            fos.close();
            doc.close();
            pkg.close();
        }
        catch(Exception e) {
            System.out.println(e.getMessage());
        }
    }
}

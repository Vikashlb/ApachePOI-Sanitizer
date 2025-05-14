package com.sanitizer.DocxSanitize;

import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFComment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

public class DocxSanitize {
    public static void main(String[] args) {
        try {
            File file = new File("TestSanitize.docx");
            //Since docx is a zip file so to unzip programmatically, Open packaging conventions(OPC) is used
            OPCPackage pkg = OPCPackage.open(file);

            //XML Word Processor Format to gain access to content and metadata in the docx
            XWPFDocument doc = new XWPFDocument(pkg);

            //Get all comments from the document
            XWPFComment[] comments = doc.getComments();

            //Iterate through each Comment's Author, Data and Initials
            if(comments != null) {
                for(XWPFComment comment : comments) {
                    comment.setAuthor(null);
                    comment.setInitials(null);
                    //comment.setDate(null);
                }
            }

            //To access the document metadata
            POIXMLProperties props = doc.getProperties();
            POIXMLProperties.CoreProperties core = props.getCoreProperties();
            core.setCreator(null);               // Clears Author
            core.setTitle(null);                 // Clears Title
            core.setSubjectProperty(null);       // Clears Subject
            core.setDescription(null);           // Clears Comments/Description
            core.setKeywords(null);              // Clears Keywords
            core.setLastModifiedByUser(null);    // Clears 'Last saved by'
            core.setCategory(null);              // Clears Category

            //Remove all extended document properties
            POIXMLProperties.ExtendedProperties ext = props.getExtendedProperties();
            ext.setCompany(null);                // Clears Company
            ext.setManager(null);                // Clears Manager
            ext.setHyperlinkBase(null);          // Clears HyperLinkBase

            //Remove all custom document properties
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

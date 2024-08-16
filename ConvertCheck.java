package org.example;

import org.apache.poi.wp.usermodel.HeaderFooterType;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.docx4j.wml.STTabJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.w3c.dom.*;
import org.w3c.dom.Document;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.*;
import java.io.File;

public class ConvertCheck {
    //private static int count=1;
    public static void main(String[] args) {

        Scanner scanner = new Scanner(System.in);
        System.out.println("Enter the directory path:");
        String directoryPath = scanner.nextLine();
        File directory = new File(directoryPath);
        File[] files = directory.listFiles((dir, name) -> name.toLowerCase().endsWith(".xml"));



        for (File file : files) {
            String xmlFilePath = file.getAbsolutePath();
            String docxFileName = file.getName().replace(".xml", ".docx");
            String docxFilePath = directoryPath + File.separator + docxFileName;

            try {
                DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
                DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
                Document doc = dBuilder.parse(new FileInputStream(xmlFilePath));

                XWPFDocument document = new XWPFDocument();
                parseElement(document, doc.getDocumentElement());
                try (FileOutputStream out = new FileOutputStream(docxFilePath)) {
                    document.write(out);
                    System.out.println("Conversion Successful.");    	        }


            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        System.out.println("\nComplete.");
    }

    private static void parseElement(XWPFDocument document, Element element) {
        NodeList nodeList = element.getChildNodes();

        for (int i = 0; i < nodeList.getLength(); i++) {
            Node node = nodeList.item(i);
            if (node instanceof Element) {
                Element elem = (Element) node;
                String tagName = elem.getTagName();
                //int count=i;

                if (tagName.equals("div") && "idinfo".equals(elem.getAttribute("type"))) {
                    createDivTypeIdInfo(document, elem);
                }
                else if (tagName.equals("div") && "toc".equals(elem.getAttribute("type"))) {
                    createTableOfContents(document,elem);
                }
                else if (tagName.equals("div") && "index".equals(elem.getAttribute("type"))) {
                    createIndex(document,elem);
                }
                else {
                    switch (tagName) {
                        case "title":
                            createTitle(document, elem);
                            break;
                        case "p":
                            createParagraph(document, elem);
                            break;
                        case "head":
                            createHeader(document, elem);
                            break;
                        case "hi":
                            createFormattedText(document.createParagraph(), elem);
                            break;
                        case "list":
                            createList(document, elem);
                            break;
                        case "item":
                            createListItem(document, elem,null);
                            break;
                        case "table":
                            createTable(document, elem);
                            break;

                        case "note":
                            createFootnoteOrEndnote(document, elem, document.createParagraph());
                            break;

                        case "amcolname":
                            createAmcolname(document,elem);
                            break;
                        case "hsep":
                            createHsep(document.createParagraph(),elem);
                            break;
                        //case "pageinfo":
                        //	handlePageInfo(document, elem);
                        //	break;
                        case "publicationstmt":
                            createPublicationStmt(document, elem);
                            break;
                        case "sourcecol":
                            createSourceCol(document, elem);
                            break;
                        case "copyright":
                            createCopyright(document, elem);
                            break;
                        case "projectdesc":
                            createProjectDescription(document, elem);
                            break;
                        case "editorialdecl":
                            createEditorialDeclaration(document,elem);
                            break;


                        case "lb":
                            handleLineBreak(document, elem);
                            break;
                        case "superscript":
                            createAnchor(document, elem, document.createParagraph());

                            // Add more cases here for different XML tags
                        default:
                            parseElement(document, elem);  // Recursively parse the element
                            break;
                    }}
            }
        }
    }





    private static void handlePageInfo(XWPFDocument document, Element pageInfoElement) {
        /*NodeList controlPgNoList = pageInfoElement.getElementsByTagName("controlpgno");
        if (controlPgNoList.getLength() > 0) {
            Element controlPgNoElement = (Element) controlPgNoList.item(0);
            Node pageNode = controlPgNoElement.getFirstChild();
            if (pageNode != null) {
                String pageNumber = pageNode.getNodeValue().trim();

                if (!pageNumber.isEmpty()) {
                    XWPFHeaderFooterPolicy headerFooterPolicy = document.getHeaderFooterPolicy();
                    if (headerFooterPolicy == null) {
                        headerFooterPolicy = document.createHeaderFooterPolicy();
                    }
                    XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
                    List<XWPFParagraph> paragraphs = footer.getParagraphs();
                    XWPFParagraph paragraph;
                    if (!paragraphs.isEmpty()) {
                        paragraph = paragraphs.get(0); // Use the first paragraph if exists
                        paragraph.removeRun(0); // Clear existing content
                    } else {
                        paragraph = footer.createParagraph(); // Create new paragraph if none exists
                    }
                    paragraph.setAlignment(ParagraphAlignment.RIGHT);
                    XWPFRun run = paragraph.createRun();

                    run.setText("Page " + pageNumber);*/

        XWPFParagraph pageBreakParagraph = document.createParagraph();
        XWPFRun pageBreakRun = pageBreakParagraph.createRun();
        pageBreakRun.addBreak(BreakType.PAGE);


        //count+=1;
    }
    //}
    //}
//    }






    private static void createDivTypeIdInfo(XWPFDocument document, Element divElement) {
        if (!"idinfo".equals(divElement.getAttribute("type"))) {
            return; // Exit if the div type is not "idinfo"
        }

        NodeList paragraphNodes = divElement.getElementsByTagName("p");

        for (int i = 0; i < paragraphNodes.getLength(); i++) {
            Element paragraphElement = (Element) paragraphNodes.item(i);
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.setAlignment(ParagraphAlignment.CENTER);

            NodeList childNodes = paragraphElement.getChildNodes();
            for (int j = 0; j < childNodes.getLength(); j++) {
                Node childNode = childNodes.item(j);
                if (childNode.getNodeType() == Node.TEXT_NODE) {
                    XWPFRun run = paragraph.createRun();
                    run.setText(childNode.getTextContent());
                    run.setBold(true); // Make text bold
                    run.setFontSize(14); // Set larger font size
                } else if (childNode.getNodeType() == Node.ELEMENT_NODE) {
                    Element childElement = (Element) childNode;
                    if (childElement.getTagName().equals("lb")) {
                        XWPFRun run = paragraph.createRun();
                        run.addBreak();
                    }
                }
            }
        }
    }


    private static void handleLineBreak(XWPFDocument document, Node lineNode) {
        Node nextNode = lineNode.getNextSibling();
        boolean skip = true;

        // Check the next four lines of the XML
        for (int i = 0; i < 4 && nextNode != null; i++) {
            if (nextNode.getNodeType() == Node.ELEMENT_NODE) {
                String tagName = nextNode.getNodeName();
                if (tagName.equals("table") || tagName.equals("head")) {
                    skip = false;
                    System.out.println("Table or header found, adding line break");
                    break;
                }
            }
            nextNode = nextNode.getNextSibling();
        }

        // If not skipping, add a line break
        if (!skip) {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.addBreak(BreakType.TEXT_WRAPPING);
            System.out.println("No table or header found, skipping line break");
        }
    }




    private static void createPublicationStmt(XWPFDocument document, Element publicationStmtElement) {
        NodeList paragraphs = publicationStmtElement.getElementsByTagName("p");

        for (int i = 0; i < paragraphs.getLength(); i++) {
            Element paragraphElement = (Element) paragraphs.item(i);

            // Create a new paragraph for each <p> element
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.setAlignment(ParagraphAlignment.CENTER);

            XWPFRun run = paragraph.createRun();
            run.setText(paragraphElement.getTextContent());
            run.setBold(true);
            // Optionally, you can apply additional formatting if needed
        }


    }


    private static void createSourceCol(XWPFDocument document, Element sourceColElement) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun run = paragraph.createRun();
        run.setText(sourceColElement.getTextContent());
        run.setBold(true);
        // Optionally, you can apply specific additional formatting if needed
    }

    private static void createCopyright(XWPFDocument document, Element copyrightElement) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun run = paragraph.createRun();
        run.setBold(true);
        run.setText(copyrightElement.getTextContent());

        // Apply specific formatting for copyright
        run.setFontSize(10); // Set font size
        run.setItalic(true); // Italicize the text
        run.setColor("808080"); // Set text color to gray (Hexadecimal color code)

        // Optionally, you can apply additional formatting if needed
    }

    private static void createProjectDescription(XWPFDocument document, Element projectDescElement) {



        NodeList paragraphs = projectDescElement.getElementsByTagName("p");

        for (int i = 0; i < paragraphs.getLength(); i++) {
            Element paragraphElement = (Element) paragraphs.item(i);

            // Create a new paragraph for each <p> element
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.setAlignment(ParagraphAlignment.CENTER);

            XWPFRun run = paragraph.createRun();
            run.setText(paragraphElement.getTextContent());
            run.setBold(true);
            run.setFontSize(14);
            // Optionally, you can apply additional formatting if needed
        }
        // Optionally, you can apply additional formatting if needed
    }

    private static void createEditorialDeclaration(XWPFDocument document, Element editorialDeclElement) {


        NodeList paragraphs = editorialDeclElement.getElementsByTagName("p");

        for (int i = 0; i < paragraphs.getLength(); i++) {
            Element paragraphElement = (Element) paragraphs.item(i);

            // Create a new paragraph for each <p> element
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.setAlignment(ParagraphAlignment.CENTER);

            XWPFRun run = paragraph.createRun();
            run.setText(paragraphElement.getTextContent());
            //run.setBold(true);
            run.setFontSize(12);
            // Optionally, you can apply additional formatting if needed
        }
    }


    private static void createHsep(XWPFParagraph paragraph, Element element) {
        // Create a run to contain the separator
        XWPFRun run = paragraph.createRun();
        // Add a horizontal rule
        run.addCarriageReturn();
        run.setText(" ");
        paragraph.setBorderBottom(Borders.SINGLE);
        run.addCarriageReturn();
    }




    private static void createAmcolname(XWPFDocument document, Element element) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(element.getTextContent());
        run.setBold(true);
        run.setItalic(true);
        paragraph.setAlignment(ParagraphAlignment.CENTER);
    }


    private static void createTitle(XWPFDocument document, Element element) {
        XWPFParagraph title = document.createParagraph();
        title.setStyle("Title");
        XWPFRun run = title.createRun();
        run.setText(element.getTextContent());
        run.setBold(true);
        run.setFontSize(20);
        title.setAlignment(ParagraphAlignment.CENTER);
    }



    private static void createParagraph(XWPFDocument document, Element element) {
        XWPFParagraph paragraph = document.createParagraph();
        handleChildNodes(paragraph, element, document);
    }

    private static void handleChildNodes(XWPFParagraph paragraph, Node node, XWPFDocument document) {
        NodeList childNodes = node.getChildNodes();
        for (int i = 0; i < childNodes.getLength(); i++) {
            Node childNode = childNodes.item(i);
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            if (childNode.getNodeType() == Node.TEXT_NODE) {
                XWPFRun run = paragraph.createRun();
                run.setText(childNode.getTextContent());
            } else if (childNode.getNodeType() == Node.ELEMENT_NODE) {
                Element childElement = (Element) childNode;
                switch (childElement.getTagName()) {
                    case "hi":
                        System.out.println("Create Formatted called");
                        createFormattedText(paragraph, childElement);
                        break;
                    case "list":
                        createList(paragraph.getDocument(), childElement);
                        break;
                    case "table":
                        createTable(paragraph.getDocument(), childElement);
                        break;
                    case "title":
                        createTitle(paragraph.getDocument(), childElement);
                        break;
                    case "head":
                        createHeader(paragraph.getDocument(), childElement);
                        break;
                    case "lb":
                        XWPFRun run = paragraph.createRun();
                        run.addBreak(BreakType.TEXT_WRAPPING);
                        break;

                    case "hsep":
                        XWPFRun rightAlignedRun = paragraph.createRun();
                        rightAlignedRun.setText(childElement.getTextContent());

                        // Add a tab and set the tab stop to the right margin
                        rightAlignedRun.addTab();
                        paragraph.setAlignment(ParagraphAlignment.LEFT);  // Keep the paragraph left-aligned
                        paragraph.setSpacingAfter(0);  // Ensure no extra space is added after hsep

                        CTPPr pPr = paragraph.getCTP().getPPr();
                        if (pPr == null) {
                            pPr = paragraph.getCTP().addNewPPr();
                        }
                        CTTabs tabs = pPr.getTabs();
                        if (tabs == null) {
                            tabs = pPr.addNewTabs();
                        }
                        CTTabStop tabStop = tabs.addNewTab();
                        //tabStop.setPos(STTabJc.RIGHT);
                        tabStop.setPos(new BigInteger("8090")); // Right margin (depends on the page width)
                        break;

                    case "item":
                        createListItem(paragraph.getDocument(), childElement, null);
                        break;
                    case "p":
                        // Nested paragraphs should be handled separately
                        XWPFParagraph nestedParagraph = paragraph.getDocument().createParagraph();

                        handleChildNodes(nestedParagraph, childElement, document);
                        break;
                    case "note":
                        createFootnoteOrEndnote(document, childElement, paragraph);
                        break;
                    case "anchor":
                        createAnchor(document, childElement, paragraph);
                        break;
                    case "superscript":
                        createAnchor(document, childElement, paragraph);
                        break;

                    default:
                        // Handle other elements or ignore them
                        break;
                }
            }
        }
    }

    private static void createFormattedText(XWPFParagraph paragraph, Element element) {
        XWPFRun run = paragraph.createRun();
        run.setText(element.getTextContent());

        switch (element.getAttribute("rend")) {
            case "smallcaps" -> {
                run.setCapitalized(true);
                run.setFontSize(8); // Adjust font size to simulate small caps
            }
            case "italics" -> run.setItalic(true);
            case "bold" -> run.setBold(true);
            case "caps" -> run.setCapitalized(true);
            default -> run.setText(element.getTextContent());
        }

        String color = element.getAttribute("color");
        if (!color.isEmpty()) {
            run.setColor(color); // Set the color from the attribute
        }
    }






    private static void createHeader(XWPFDocument document, Element element) {
        handlePageInfo(document, element);
        XWPFParagraph header = document.createParagraph();
        header.setStyle("Heading1");
        XWPFRun run = header.createRun();
        run.setText(element.getTextContent());
        run.setBold(true);
        run.setFontSize(16);

        String color = element.getAttribute("color");
        if (!color.isEmpty()) {
            run.setColor(color); // Set the color from the attribute
        }

    }



    private static void createFormattedText(XWPFDocument document, Element element) {
        XWPFParagraph paragraph = document.createParagraph();


        XWPFRun run = paragraph.createRun();
        run.setText(element.getTextContent());
        switch (element.getAttribute("rend")) {
            case "smallcaps" -> {
                run.setCapitalized(true);
                run.setFontSize(8); // Adjust font size to simulate small caps
            }
            case "italics" -> {
                run.setItalic(true);
            }
            case "bold" -> run.setBold(true);
        }

        // Add more formatting as needed
    }

    private static void createList(XWPFDocument document, Element element) {

        NodeList items = element.getElementsByTagName("item");
        String listType = element.getAttribute("type");
        BigInteger numID = null;

        if ("ordered".equalsIgnoreCase(listType)) {
            numID = addOrderedList(document);
        } else if ("unordered".equalsIgnoreCase(listType)) {
            numID = addUnorderedList(document);
        }

        for (int i = 0; i < items.getLength(); i++) {
            Element item = (Element) items.item(i);
            createListItem(document, item, numID);
        }
    }

    private static BigInteger addOrderedList(XWPFDocument document) {
        XWPFNumbering numbering = document.createNumbering();
        CTAbstractNum ctAbstractNum = CTAbstractNum.Factory.newInstance();
        ctAbstractNum.setAbstractNumId(BigInteger.valueOf(0));
        CTLvl ctLvl = ctAbstractNum.addNewLvl();
        ctLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
        ctLvl.addNewLvlText().setVal("%1.");
        BigInteger abstractNumId = numbering.addAbstractNum(new XWPFAbstractNum(ctAbstractNum));
        return numbering.addNum(abstractNumId);
    }

    private static BigInteger addUnorderedList(XWPFDocument document) {
        XWPFNumbering numbering = document.createNumbering();
        CTAbstractNum ctAbstractNum = CTAbstractNum.Factory.newInstance();
        ctAbstractNum.setAbstractNumId(BigInteger.valueOf(1));
        CTLvl ctLvl = ctAbstractNum.addNewLvl();
        ctLvl.addNewNumFmt().setVal(STNumberFormat.BULLET);
        ctLvl.addNewLvlText().setVal("•");
        BigInteger abstractNumId = numbering.addAbstractNum(new XWPFAbstractNum(ctAbstractNum));
        return numbering.addNum(abstractNumId);
    }



    private static void createListItem(XWPFDocument document, Element itemElement, BigInteger numID) {
        XWPFParagraph paragraph = document.createParagraph();
        if (numID != null) {
            paragraph.setNumID(numID);
        }

        NodeList childNodes = itemElement.getChildNodes();
        for (int i = 0; i < childNodes.getLength(); i++) {
            Node childNode = childNodes.item(i);
            if (childNode.getNodeType() == Node.TEXT_NODE) {
                XWPFRun run = paragraph.createRun();
                run.setText(childNode.getTextContent());
            } else if (childNode.getNodeType() == Node.ELEMENT_NODE) {
                Element childElement = (Element) childNode;
                switch (childElement.getTagName()) {

                    case "list":
                        createList(document, childElement);
                        break;
                    case "table":
                        createTable(document, childElement);
                        break;
                    case "hi":
                        createFormattedText(document, childElement);
                        break;
                    case "title":
                        createTitle(document, childElement);
                        break;
                    case "head":
                        createHeader(document, childElement);
                        break;
                    case "lb":
                        handleLineBreak(document, childElement);
                        break;
                    case "item":
                        createListItem(document, childElement,null);
                        break;
                    case "note":
                        createFootnoteOrEndnote(document, childElement, paragraph);
                        break;
                    case "anchor":
                        createAnchor(document, childElement, paragraph);
                        break;
                    case "p":
                        createParagraph(document, childElement);
                        break;

                    default:
                        break;
                }
            }
        }
    }

    @SuppressWarnings("deprecation")
    private static void createTable(XWPFDocument document, Element tableElement) {

        NodeList tableTextNodes = tableElement.getElementsByTagName("tabletext");
        for (int i = 0; i < tableTextNodes.getLength(); i++) {
            Element tableTextElement = (Element) tableTextNodes.item(i);
            NodeList cellNodes = tableTextElement.getElementsByTagName("cell");

            int numCols = 4; // Number of columns in the table
            // Calculate the number of rows and columns
            int numRows = cellNodes.getLength() / 4; // Assuming 4 cells per row


            // Create a table in the document
            XWPFTable table = document.createTable(numRows, numCols);

            // Set table width to 100% of the page width
            CTTbl tableCT = table.getCTTbl();
            CTTblPr tablePr = tableCT.getTblPr() == null ? tableCT.addNewTblPr() : tableCT.getTblPr();
            CTTblWidth tblW = tablePr.isSetTblW() ? tablePr.getTblW() : tablePr.addNewTblW();
            tblW.setW(BigInteger.valueOf(50000)); // You can adjust this value for better fitting

            // Fill the table with data from <cell> elements
            int cellIndex = 0;
            for (int rowIndex = 0; rowIndex < numRows; rowIndex++) {
                XWPFTableRow row = table.getRow(rowIndex);
                for (int colIndex = 0; colIndex < numCols; colIndex++) {
                    XWPFTableCell cell = row.getCell(colIndex);
                    if (cellIndex < cellNodes.getLength()) {
                        Element cellElement = (Element) cellNodes.item(cellIndex);
                        XWPFParagraph paragraph = cell.getParagraphs().get(0);
                        XWPFRun run = paragraph.createRun();
                        run.setText(cellElement.getTextContent());

                        if (rowIndex == 0) { // Format headers
                            run.setBold(true);
                            run.setFontSize(14);
                            cell.setColor("CCCCCC"); // Highlight color (light gray)
                        } else {
                            run.setFontSize(12); // Regular text size for other cells
                        }

                        // Set cell width
                        CTTcPr tcPr = cell.getCTTc().addNewTcPr();
                        CTTblWidth cellWidth = tcPr.addNewTcW();
                        cellWidth.setType(STTblWidth.DXA);
                        cellWidth.setW(BigInteger.valueOf(12500)); // Adjust cell width as needed

                        // Ensure text wrapping
                        cell.getParagraphs().get(0).setWordWrap(true);

                        cellIndex++;
                    }
                }
            }
        }
    }





    private static void createFootnoteOrEndnote(XWPFDocument document, Element noteElement, XWPFParagraph paragraph) {
        // Create a new footnote in the document
        XWPFFootnote footnote = document.createFootnote();
        XWPFParagraph footnoteParagraph = footnote.createParagraph();

        // Iterate over the child nodes of the note element
        NodeList childNodes = noteElement.getChildNodes();
        for (int i = 0; i < childNodes.getLength(); i++) {
            Node node = childNodes.item(i);
            XWPFRun run = footnoteParagraph.createRun();
            if (node instanceof Element) {
                Element childElement = (Element) node;
                if (childElement.getTagName().equals("superscript")) {
                    run.setText(childElement.getTextContent());
                    run.setSubscript(VerticalAlign.SUPERSCRIPT);
                    run.setFontSize(7);

                } else {
                    run.setText(childElement.getTextContent());
                    run.setFontSize(7);

                }
            } else if (node instanceof Text) {
                run.setText(node.getTextContent());
                run.setFontSize(7);

            }
        }

        // Add a reference to the footnote in the main document paragraph
       CTFtnEdnRef ref = paragraph.getCTP().addNewR().addNewFootnoteReference();
        ref.setId(footnote.getCTFtnEdn().getId());
    }








    private static void createTableOfContents(XWPFDocument document, Element element) {
        NodeList childNodes = element.getChildNodes();
        for (int i = 0; i < childNodes.getLength(); i++) {
            Node childNode = childNodes.item(i);
            if (childNode.getNodeType() == Node.ELEMENT_NODE) {
                Element childElement = (Element) childNode;
                switch (childElement.getTagName()) {
                    case "head":
                        createHeader(document, childElement);
                        break;
                    case "list":
                        createList(document, childElement);
                        break;
                    case "p":
                        createTOCEntry(document, childElement);
                        break;
                    default:
                        break;
                }
            }
        }
    }

    private static void createTOCEntry(XWPFDocument document, Element element) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setStyle("TOCHeading"); // Set style for TOC entries
        XWPFRun run = paragraph.createRun();
        run.setText(element.getTextContent());
    }


    private static void createIndex(XWPFDocument document, Element element) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        //run.setText("Index");
        run.setBold(true);
        createTableOfContents(document, element);
        // Implement Index creation logic
    }

    private static void createAnchor(XWPFDocument document, Element anchorElement, XWPFParagraph paragraph) {

        XWPFRun run = paragraph.createRun();

        // Set the content of the anchor element in superscript
        run.setText(anchorElement.getTextContent());
        run.setFontSize(12); // Adjust font size if needed
        run.setSubscript(VerticalAlign.SUPERSCRIPT);
        run.setColor("FF0000");
    }


    // Add more helper methods here for other XML tags and their DOCX counterparts
}



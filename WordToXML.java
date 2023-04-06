import java.io.*;
import java.math.BigInteger;
import java.util.List;

import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.drawingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;

public class WordToXML {
  /**
 * Reads a Word document consisting of texts, images, tables, header and footer, and
 * stores them in an XML file.
 *
 * @param args Command line arguments.
 * @throws IOException If an I/O error occurs while reading or writing the files.
 * @throws XMLStreamException If an error occurs while writing the XML file.
 */
    public static void main(String[] args) throws XMLStreamException, IOException,NullPointerException {
        System.out.println("Process Started.....");
        String inputFilePath = "input.docx";
        String outputFilePath = "output";

        InputStream inputStream = new FileInputStream(inputFilePath);


             XWPFDocument document = new XWPFDocument(inputStream);
             BufferedWriter writer = new BufferedWriter(new FileWriter(outputFilePath));
             try {

            writer.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            writer.write("<document>");

            XMLOutputFactory factory = XMLOutputFactory.newInstance();
            XMLStreamWriter xmlWriter = factory.createXMLStreamWriter(writer);

            for (IBodyElement element : document.getBodyElements()) {
                if (element instanceof XWPFParagraph) {
                    //extra
                    
                    
                    //extra

                    writeParagraphContent((XWPFParagraph) element, xmlWriter);
                } else if (element instanceof XWPFTable) {
                    writeTableContent((XWPFTable) element, writer);
                }
            }

            writer.write("</document>");
            writer.close();
            System.out.println("File converted to XML Successfully as "+ outputFilePath+".xml");
        } catch (IOException e) {
            System.out.println("Error reading Word document: " + e.getMessage());
        }
    }

    private static void writeParagraphContent(XWPFParagraph paragraph, XMLStreamWriter writer) throws XMLStreamException {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs != null && runs.size() > 0) {
            writer.writeStartElement("p");
            writer.writeAttribute("align", paragraph.getAlignment().toString());
    
            for (XWPFRun run : runs) {
                if (run.getEmbeddedPictures().size() > 0){
                    for (XWPFPicture picture : run.getEmbeddedPictures()) {
                            if ((picture != null && picture.getCTPicture() != null && picture.getCTPicture().getBlipFill() != null && picture.getCTPicture().getBlipFill().getBlip() != null)) {
                                writer.writeStartElement("img");
                                writer.writeAttribute("src", "data:image/png;base64," + picture.getCTPicture().getBlipFill().getBlip().getEmbed());
                                writer.writeEndElement();
                            }
                        
                    }
                
                    
                }
                writer.writeStartElement("r");
    
                // Write text content
                String text = run.getText(0);
                if (text != null && !text.isEmpty()) {
                    writer.writeStartElement("t");
                    writer.writeCharacters(text);
                    writer.writeEndElement();
                }
    
                // Write character properties
                CTRPr charProps = run.getCTR().getRPr();
               // System.out.println("Got a non empty charProps");
                if (charProps != null) {
                    // Font family
                    CTFonts fonts = charProps.getRFonts();
                    
                    if (fonts != null) {
                       // System.out.println("Got a non empty font");
                       // System.out.println(fonts);
                        String fontFamily = fonts.getAscii();
                        if (fontFamily != null && !fontFamily.isEmpty()) {
                            //System.out.println("Non empty fontfamily");
                            System.out.println("Non empty font family");
                            System.out.println(fontFamily);
                            writer.writeStartElement("fontFamily");
                            writer.writeCharacters(fontFamily);
                            writer.writeEndElement();
                        }
                    }else{
                       // System.out.println("Null font");
                    }
    
                    // Font size
                    Object szObj = charProps.getSz();
                    if (szObj instanceof BigInteger) {
                       // System.out.println("Big int");
                        BigInteger fontSize = (BigInteger) szObj;
                    if (fontSize != null) {
                        System.out.println("font size"+ fontSize.toString());
                        String fontSizeStr = fontSize.toString();
                        if (fontSizeStr != null && !fontSizeStr.isEmpty()) {
                            writer.writeStartElement("fontSize");
                            writer.writeCharacters(fontSizeStr);
                            writer.writeEndElement();
                        }
                    }
                }
    
                    // Font color
                    CTColor color = (CTColor) charProps.getColor();

                    if (color != null) {
                        System.out.println("color is "+ color);
                        String colorStr = color.toString();
                       // System.out.println("Got a non empty color");

                        if (colorStr != null && !colorStr.isEmpty()) {
                            writer.writeStartElement("color");
                            writer.writeCharacters(colorStr);
                            writer.writeEndElement();
                        }
                    }else{
                        //System.out.println("no color");
                    }
    
                    // Bold
                    boolean isBold = charProps.isSetB() && charProps.getB().getVal() == STOnOff.TRUE;
                    if (isBold) {
                        writer.writeStartElement("b");
                        writer.writeEndElement();
                    }
    
                    // Italic
                    boolean isItalic = charProps.isSetI() && charProps.getI().getVal() == STOnOff.TRUE;
                    if (isItalic) {
                        writer.writeStartElement("i");
                        writer.writeEndElement();
                    }
    
                    // Underline
                    boolean isUnderline = charProps.isSetU() && charProps.getU().getVal() != STUnderline.NONE;
                    if (isUnderline) {
                        //System.out.println("underlined found ");
                        writer.writeStartElement("u");
                        writer.writeEndElement();
                    }
                }
    
                writer.writeEndElement(); // end r
            }
    
            writer.writeEndElement(); // end p
        }
    }
    

    private static void writeTableContent(XWPFTable table, BufferedWriter writer) throws IOException, XMLStreamException {
        writer.write("<table>");
        for (XWPFTableRow row : table.getRows()) {
            writer.write("<row>");
            for (XWPFTableCell cell : row.getTableCells()) {
                writer.write("<cell>");
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                    XMLOutputFactory factory = XMLOutputFactory.newInstance();
            XMLStreamWriter xmlWriter = factory.createXMLStreamWriter(writer);
                    writeParagraphContent(paragraph, xmlWriter);
                }
                writer.write("</cell>");
            }
            writer.write("</row>");
        }
        writer.write("</table>");
        
    }
}

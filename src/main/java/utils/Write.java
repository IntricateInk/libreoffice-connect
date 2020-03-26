
// Write.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2015

/* A growing collection of utility functions to make Office
   easier to use. They are currently divided into the following
   groups:

     * text doc methods
     * model/view cursor methods
     * text cursor property methods
     * text writing methods
     * extract text from document
     * text property methods
     * style methods
     * headers and footers
     * adding elements
         - formulae, bookmark, text frame, table
     * adding image methods 
     * extracting graphics
     * linguistic API.

*/

package utils;

import java.io.*;
import java.util.*;
import java.awt.image.*;
import javax.imageio.*;

import com.sun.star.beans.*;
import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.text.*;
import com.sun.star.uno.*;
import com.sun.star.awt.*;
import com.sun.star.util.*;
import com.sun.star.document.*;
import com.sun.star.container.*;
import com.sun.star.text.*;
import com.sun.star.style.*;
import com.sun.star.view.*;
import com.sun.star.drawing.*;
import com.sun.star.table.*;
import com.sun.star.graphic.*;
import com.sun.star.linguistic2.*;
import com.sun.star.deployment.*;
import com.sun.star.resource.*;

import com.sun.star.uno.Exception;
import com.sun.star.io.IOException;


public class Write
{


  // ---------------------------- text doc methods ----------------------------

  public static XTextDocument openDoc(String fnm, XComponentLoader loader)
  {
    XComponent doc = Lo.openDoc(fnm, loader);
    if (doc == null) {
      System.out.println("Document is null");
      return null;
    }

    if (!isText(doc)) {
      System.out.println("Not a text document; closing " + fnm);
      Lo.closeDoc(doc);
      return null;
    }

    XTextDocument textDoc = Lo.qi(XTextDocument.class, doc);
    if (textDoc == null) {
      System.out.println("Not a text document; closing " + fnm);
      Lo.closeDoc(doc);
      return null;
    }

    return textDoc;
  }  // end of openDoc()




  public static boolean isText(XComponent doc)
  // is doc a text document?
  {  return Info.isDocType(doc, Lo.WRITER_SERVICE);  }



  public static XTextDocument getTextDoc(XComponent doc)
  // cast doc into a text document (if possible)
  {
    if (doc == null) {
      System.out.println("Document is null");
      return null;
    }
    XTextDocument textDoc = Lo.qi(XTextDocument.class, doc);
    if (textDoc == null)
      System.out.println("Not a text document");
    return textDoc;
  }  // end of getTextDoc()




  public static XTextDocument createDoc(XComponentLoader loader)
  { XComponent doc = Lo.createDoc(Lo.WRITER_STR, loader);  
    return Lo.qi(XTextDocument.class, doc);
  }


  public static XTextDocument createDocFromTemplate(String templatePath,
                                                    XComponentLoader loader)
  { XComponent doc = Lo.createDocFromTemplate(templatePath, loader);  
    return Lo.qi(XTextDocument.class, doc);
  }


  public static void closeDoc(XTextDocument textDoc)
  { XCloseable closeable = Lo.qi(XCloseable.class, textDoc);
    Lo.close(closeable);
  }


  public static void saveDoc(XTextDocument textDoc, String fnm)
  { XComponent doc = Lo.qi(XComponent.class, textDoc);
    Lo.saveDoc(doc, fnm);
  }



  public static XTextDocument openFlatDocUsingTextTemplate(String fnm, 
                                                 String templatePath,
                                                 XComponentLoader loader)
  // open a new text document applying the template as formatting to the flat XML file
  {
    // check the filename
    if (fnm == null) {
      System.out.println("Filename is null");
      return null;
    }

    String openFileURL = null;
    if (!FileIO.isOpenable(fnm)) {
      if (Lo.isURL(fnm)) {
        System.out.println("Will treat filename as a URL: \"" + fnm + "\"");
        openFileURL = fnm;
      }
      else
        return null;
    }
    else {
      openFileURL = FileIO.fnmToURL(fnm);
      if (openFileURL == null)
       return null;
    }

    String templateExt = Info.getExt(templatePath);
    if (!templateExt.equals("ott")) {
      System.out.println("Can only apply a text template as formatting");
      return null;
    }

    // create a new document using the template
    XComponent doc = Lo.createDocFromTemplate(templatePath, loader);
    if (doc == null)
      return null;

    XTextDocument textDoc = Lo.qi(XTextDocument.class, doc);
    XTextCursor cursor = getCursor(textDoc);
    try {
      cursor.gotoEnd(true);
      XDocumentInsertable di = Lo.qi(XDocumentInsertable.class, cursor);
                  // XDocumentInsertable only works with text files
      if (di == null)
        System.out.println("Document inserter could not be created");
      else
        di.insertDocumentFromURL(openFileURL, new PropertyValue[0]);
             // Props.makeProps("FilterName", "OpenDocument Text Flat XML"));
             // these props do not work
    }
    catch (java.lang.Exception e)
    {  System.out.println("Could not insert document"); }

     return textDoc;
  }  // end of openFlatDocUsingTextTemplate()


  // --------------------- model cursor methods -------------------------------



  public static XTextCursor getCursor(XTextDocument textDoc)
  // get model cursor from a text document
  {
    XText xText = textDoc.getText();
    if (xText == null) {
      System.out.println("Text not found in document");
      return null;
    }
    else
      return xText.createTextCursor();
  }  // end of getCursor()



  public static XPropertySet getTextCursorProps(XTextDocument textDoc)
  { XTextCursor cursor = getCursor(textDoc);
    return Lo.qi(XPropertySet.class, cursor);
  }  // end of getTextCursorProps()





  public static XWordCursor getWordCursor(XTextDocument textDoc)
  {  
    XTextCursor cursor = getCursor(textDoc);
    if (cursor == null) {
      System.out.println("Text cursor is null");
      return null;
    }
    else
      return Lo.qi(XWordCursor.class, cursor);  
  }  // end of getWordCursor()



  public static XSentenceCursor getSentenceCursor(XTextDocument textDoc)
  {  
    XTextCursor cursor = getCursor(textDoc);
    if (cursor == null) {
      System.out.println("Text cursor is null");
      return null;
    }
    else
      return Lo.qi(XSentenceCursor.class, cursor);  
  }  // end of getSentenceCursor()



  public static XParagraphCursor getParagraphCursor(XTextDocument textDoc)
  {  
    XTextCursor cursor = getCursor(textDoc);
    if (cursor == null) {
      System.out.println("Text cursor is null");
      return null;
    }
    else
      return Lo.qi(XParagraphCursor.class, cursor);  
  }  // end of getParagraphCursor()



  public static int getPosition(XTextCursor cursor)
  {  return (cursor.getText().getString()).length();  }



  // ---------------------- view cursor methods ------------------------------


  public static XTextViewCursor getViewCursor(XTextDocument textDoc)
  {
    XModel model = Lo.qi(XModel.class, textDoc);
    XController xController = model.getCurrentController();
    
    // the controller gives us the TextViewCursor
    XTextViewCursorSupplier supplier = Lo.qi(XTextViewCursorSupplier.class, xController);
    return supplier.getViewCursor();
  }  // end of getViewCursor()



  public static XTextCursor getCursor(XTextViewCursor tvCursor)
  {  return tvCursor.getText().createTextCursorByRange(tvCursor);   } 




  public static int getCurrentPage(XTextViewCursor tvCursor)
  {
    XPageCursor pageCursor = Lo.qi(XPageCursor.class, tvCursor);
    if (pageCursor == null) {
      System.out.println("Could not create a page cursor");
      return -1;
    }
    // System.out.println("current page = " + pageCursor.getPage());
    return pageCursor.getPage();
  }  // end of getCurrentPage()



  public static String getCoordStr(XTextViewCursor tvCursor)
  {  return "(" + tvCursor.getPosition().X + ", " + tvCursor.getPosition().Y + ")"; }




  public static int getPageCount(XTextDocument textDoc)
  {
    XModel model = Lo.qi(XModel.class, textDoc);
    XController xController = model.getCurrentController();
    return ((Integer)Props.getProperty(xController, "PageCount")).intValue();
  }  // end of getPageCount()




  public static XPropertySet getTextViewCursorPropSet(XTextDocument textDoc)
  { XTextViewCursor xViewCursor = getViewCursor(textDoc);
    return Lo.qi(XPropertySet.class, xViewCursor);
  } 




  // ------------------------- text writing methods ------------------------------------


  public static int append(XTextCursor cursor, String text)
  { cursor.setString(text);
    cursor.gotoEnd(false);
    return getPosition(cursor);
  }  // end of append()



  public static int append(XTextCursor cursor, short ctrlChar)
  // add control character to end of document
  {
    XText xText = cursor.getText();
    xText.insertControlCharacter(cursor, ctrlChar, false);
    cursor.gotoEnd(false);
    return getPosition(cursor);
  }



  public static int append(XTextCursor cursor, XTextContent textContent)
  // embed text content (e.g. table, text field) at end of document
  {
    XText xText = cursor.getText();
    xText.insertTextContent(cursor, textContent, false);
    cursor.gotoEnd(false);
    return getPosition(cursor);
  }


  public static int appendDateTime(XTextCursor cursor)
  // append two DateTime fields, one for the date, one for the time
  { 
    XTextField dtField = Lo.createInstanceMSF(XTextField.class, 
                                   "com.sun.star.text.TextField.DateTime");

    Props.setProperty(dtField, "IsDate", true);    // so date is reported
    append(cursor, dtField);
    append(cursor, "; ");

    dtField = Lo.createInstanceMSF(XTextField.class, 
                                   "com.sun.star.text.TextField.DateTime");
    Props.setProperty(dtField, "IsDate", false);    // so time is reported
    return append(cursor, dtField);
  }  // end of appendDateTime()



  public static int appendPara(XTextCursor cursor, String text)
  { append(cursor, text);
    append(cursor, ControlCharacter.PARAGRAPH_BREAK);
    return getPosition(cursor);
  }



  public static void endLine(XTextCursor cursor)
  {  append(cursor, ControlCharacter.LINE_BREAK);  }



  public static void endParagraph(XTextCursor cursor)
  {  append(cursor, ControlCharacter.PARAGRAPH_BREAK);  }


  public static void pageBreak(XTextCursor cursor)
  { Props.setProperty(cursor, "BreakType", BreakType.PAGE_AFTER);
    endParagraph(cursor);
  } 


  public static void columnBreak(XTextCursor cursor)
  { Props.setProperty(cursor, "BreakType", BreakType.COLUMN_AFTER);
    endParagraph(cursor);
  } 



  public static void insertPara(XTextCursor cursor, String para, String paraStyle)
  {
    XText xText = cursor.getText();
    xText.insertString(cursor, para, false);
    xText.insertControlCharacter(cursor, ControlCharacter.PARAGRAPH_BREAK, false);
    stylePrevParagraph(cursor, paraStyle);
  }



  // --------------------- extract text from document ------------------------


  public static String getAllText(XTextCursor cursor)
  // return the text part of the document
  {
    cursor.gotoStart(false);
    cursor.gotoEnd(true);
    String text = cursor.getString();
    cursor.gotoEnd(false);    // to deselect everything
    return text;
  }  // end of getAllText()



  public static XEnumeration getEnumeration(Object obj)
  {
    XEnumerationAccess enumAccess = Lo.qi(XEnumerationAccess.class, obj);
    if (enumAccess == null) {
      System.out.println("Could not create enumeration");
      return null;
    }
    else
      return enumAccess.createEnumeration();
  }  // end of getEnumeration()


  // ------------------------ text cursor property methods -----------------------------------


  public static void styleLeftBold(XTextCursor cursor, int pos)
  {  styleLeft(cursor, pos, "CharWeight", com.sun.star.awt.FontWeight.BOLD);  }


  public static void styleLeftItalic(XTextCursor cursor, int pos)
  {  styleLeft(cursor, pos, "CharPosture", com.sun.star.awt.FontSlant.ITALIC);  }


  public static void styleLeftColor(XTextCursor cursor, int pos, java.awt.Color col)
  {  styleLeft(cursor, pos, "CharColor", Lo.getColorInt(col));  }


  public static void styleLeftCode(XTextCursor cursor, int pos)
  {  styleLeft(cursor, pos, "CharFontName", "Courier New");
     styleLeft(cursor, pos, "CharHeight", 10);
  }



  public static void styleLeft(XTextCursor cursor, 
                                 int pos, String propName, Object propVal)
  {
    Object oldValue = Props.getProperty(cursor, propName);

    int currPos = getPosition(cursor);
    cursor.goLeft((short)(currPos-pos), true);
    Props.setProperty(cursor, propName, propVal);

    cursor.goRight((short)(currPos-pos), false);
    Props.setProperty(cursor, propName, oldValue);
  }  // end of styleLeft()


/*
  public static void stylePrevWord(XTextCursor cursor,
                                                String propName, Object propVal)
  {
    Object oldValue = Props.getProperty(cursor, propName);

    XWordCursor wordCursor = Lo.qi(XWordCursor.class, cursor); 
    wordCursor.gotoPreviousWord(true);   // select previous word
    Props.setProperty(wordCursor, propName, propVal);

    // reset 
    wordCursor.gotoEndOfWord(false); 
    Props.setProperty(cursor, propName, oldValue);
  }  // end of stylePrevWord()
*/



  public static void stylePrevParagraph(XTextCursor cursor, Object propVal)
  {  stylePrevParagraph(cursor, "ParaStyleName", propVal);  }


  public static void stylePrevParagraph(XTextCursor cursor,
                                                String propName, Object propVal)
  {
    Object oldValue = Props.getProperty(cursor, propName);

    XParagraphCursor paraCursor = Lo.qi(XParagraphCursor.class, cursor);
    paraCursor.gotoPreviousParagraph(true);   // select previous paragraph
    Props.setProperty(paraCursor, propName, propVal);

    // reset 
    paraCursor.gotoNextParagraph(false);
    Props.setProperty(cursor, propName, oldValue);
  }  // end of stylePrevParagraph()




  // ---------------------------- style methods -------------------------------


  //public static String[] getParaStyleNames(XTextDocument textDoc)
  //{  return Info.getStyleNames(textDoc, "ParagraphStyles");  } 


  //public static XNameContainer getParaStyles(XTextDocument textDoc)
  //{  return  Info.getStyleContainer(textDoc, "ParagraphStyles");  }


  //public static XNameContainer getPageStyles(XTextDocument textDoc)
  //{  return  Info.getStyleContainer(textDoc, "PageStyles");  }




  public static int getPageTextWidth(XTextDocument textDoc)
  // get the width of the page's text area
  {
    XPropertySet props = Info.getStyleProps(textDoc,"PageStyles", "Standard");
    if (props == null) {
      System.out.println("Could not access the standard page style");
      return 0;
    }

    try {
      int width = ((Integer)props.getPropertyValue("Width")).intValue();
      int leftMargin = ((Integer)props.getPropertyValue("LeftMargin")).intValue();
      int rightMargin = ((Integer)props.getPropertyValue("RightMargin")).intValue();
      return (width - (leftMargin + rightMargin));
    }
    catch (Exception e) 
    {  System.out.println("Could not access standard page style dimensions: " + e); 
       return 0;
    }
  }  // end of getPageSize()



  public static Size getPageSize(XTextDocument textDoc)
  // get the size of the page
  {
    XPropertySet props = Info.getStyleProps(textDoc,"PageStyles", "Standard");
    if (props == null) {
      System.out.println("Could not access the standard page style");
      return null;
    }

    try {
      int width = ((Integer)props.getPropertyValue("Width")).intValue();
      int height = ((Integer)props.getPropertyValue("Height")).intValue();
      // System.out.println("Size: " + width + ", " + height);  // in 1/100 mm units
            /* probably letter, which is 8.5 x 11 inches ==
               215.9 x 279.4 mm ==  21590 x 27940 units
            */
      return new Size(width, height);
    }
    catch (Exception e) 
    {  System.out.println("Could not access standard page style dimensions: " + e); 
       return null;
    }
  }  // end of getPageSize()



  public static void setA4PageFormat(XTextDocument textDoc)
  // set the page format to A4
  {
    XPrintable xPrintable = Lo.qi(XPrintable.class, textDoc);
    PropertyValue[] printerDesc = new PropertyValue[1];
      
    // Paper Format           
    printerDesc[0] = new PropertyValue();
    printerDesc[0].Name = "PaperFormat";
    printerDesc[0].Value = PaperFormat.A4;

  /*
    // Paper Orientation           
    printerDesc[1] = new PropertyValue();
    printerDesc[1].Name = "PaperOrientation";
    printerDesc[1].Value = PaperOrientation.LANDSCAPE;
  */
      
    xPrintable.setPrinter(printerDesc);
  }  // end of setA4PageFormat()



  // ------------------------ headers and footers ----------------------------


  public static void setPageNumbers(XTextDocument textDoc)
  /* Modify the footer via the page style for the document. 
     Put page number & count in the center of the footer in Times New Roman, 12pt
  */
  {
    XPropertySet props = Info.getStyleProps(textDoc,"PageStyles", "Standard");
    if (props == null) {
      System.out.println("Could not access the standard page style container");
      return;
    }

    try {
      props.setPropertyValue("FooterIsOn", Boolean.TRUE);
                                 // Footer must be turned on in the document

      XText footerText = Lo.qi(XText.class, props.getPropertyValue("FooterText"));
      XTextCursor footerCursor = footerText.createTextCursor();
     
      // set footer text properties
      Props.setProperty(footerCursor, "CharFontName", "Times New Roman");
      // Props.setProperty(footerCursor, "CharFontStyleName", "Regular");
      Props.setProperty(footerCursor, "CharHeight", 12.0f);
      Props.setProperty(footerCursor, "ParaAdjust", ParagraphAdjust.CENTER);

      // add text fields to the footer
      append(footerCursor, getPageNumber());
      append(footerCursor, " of ");
      append(footerCursor, getPageCount());
    } 
    catch (Exception ex) 
    {  System.out.println(ex); }
  }  // end of setPageNumbers()


  public static XTextField getPageNumber()
  // return arabic style number showing current page value
  {
    XTextField numField = Lo.createInstanceMSF(XTextField.class, 
                                          "com.sun.star.text.TextField.PageNumber");
    Props.setProperty(numField, "NumberingType", NumberingType.ARABIC);
    Props.setProperty(numField, "SubType", PageNumberType.CURRENT);
    return numField;
  }


  public static XTextField getPageCount()
  // return arabic style number showing current page count
  {
    XTextField pcField = Lo.createInstanceMSF(XTextField.class, 
                                            "com.sun.star.text.TextField.PageCount");
    Props.setProperty(pcField, "NumberingType", NumberingType.ARABIC);
    return pcField;
  }




  public static void setHeader(XTextDocument textDoc, String hText)
  /* Modify the header via the page style for the document. 
     Put the text on the right hand side in the header in
     Times New Roman, 10pt */
  {
    XPropertySet props = Info.getStyleProps(textDoc,"PageStyles", "Standard");
    if (props == null) {
      System.out.println("Could not access the standard page style container");
      return;
    }

    try {
      props.setPropertyValue("HeaderIsOn", true);
                                 // header must be turned on in the document
      // props.setPropertyValue("TopMargin", 2200);

      // access the header XText and its properties
      XText headerText = Lo.qi(XText.class, props.getPropertyValue("HeaderText"));
      XTextCursor headerCursor = headerText.createTextCursor();
      headerCursor.gotoEnd(false);
     
      XPropertySet headerProps = Lo.qi(XPropertySet.class, headerCursor);
      headerProps.setPropertyValue("CharFontName", "Times New Roman");
      headerProps.setPropertyValue("CharHeight", 10);
      headerProps.setPropertyValue("ParaAdjust", ParagraphAdjust.RIGHT);

      headerText.setString(hText + "\n");
    } 
    catch (Exception ex) 
    {  System.out.println(ex); }
  }  // end of setHeader()




  public static XDrawPage getDrawPage(XTextDocument doc)
  {
    XDrawPageSupplier xSuppPage = Lo.qi(XDrawPageSupplier.class, doc);
    if (xSuppPage != null)
      return xSuppPage.getDrawPage();
    else {
      System.out.println("No draw page found");
      return null;
    }
  }  // end of getDrawPage()



  // -------------------------- adding elements ----------------------------



  public static void addFormula(XTextCursor cursor, String formula) 
  {
    try {
      XTextContent embedContent = Lo.createInstanceMSF(XTextContent.class, 
                                                 "com.sun.star.text.TextEmbeddedObject");
      if (embedContent == null) {
        System.out.println("Could not create a formula embedded object");
        return;
      }

      // set class ID for type of object being inserted
      XPropertySet props = Lo.qi(XPropertySet.class, embedContent);
      props.setPropertyValue("CLSID", Lo.MATH_CLSID);  // a formula
      props.setPropertyValue("AnchorType", TextContentAnchorType.AS_CHARACTER);

      // insert object in document
      append(cursor, embedContent);
      endLine(cursor);

      // access object's model
      XEmbeddedObjectSupplier2 embedObjSupplier = 
                Lo.qi(XEmbeddedObjectSupplier2.class, embedContent);
      XComponent embedObjModel = embedObjSupplier.getEmbeddedObject();

      // insert formula into the object
      XPropertySet formulaProps = Lo.qi(XPropertySet.class, embedObjModel);
      formulaProps.setPropertyValue("Formula", formula);
      System.out.println("Inserted formula \"" + formula + "\"");
    }
    catch (Exception e) {
      System.out.println("Insertion of formula \"" + formula + "\" failed: " + e);
    }
  }  // end of addFormula()


  public static void addHyperlink(XTextCursor cursor, String label, String urlStr)
  {
    XTextContent link =  Lo.createInstanceMSF(XTextContent.class,
                                              "com.sun.star.text.TextField.URL");
    if (link == null) {
      System.out.println("Could not create a hyperlink");
      return;
    }

    // create a hyperlink
    Props.setProperty(link, "URL", urlStr);
    Props.setProperty(link, "Representation", label);

    append(cursor, link);
  }  // end of addHyperlink()



  public static void addBookmark(XTextCursor cursor, String name)
  {
    XTextContent bmkContent = Lo.createInstanceMSF(XTextContent.class, 
                                               "com.sun.star.text.Bookmark");
    if (bmkContent == null) {
      System.out.println("Could not create a bookmark");
      return;
    }

    XNamed bmkNamed = Lo.qi(XNamed.class, bmkContent);
    bmkNamed.setName(name);

    append(cursor, bmkContent);
  }  // end of addBookmark()


  public static XTextContent findBookmark(XTextDocument doc, String bmName)
  {
    XBookmarksSupplier supplier = Lo.qi(XBookmarksSupplier.class, doc);
    if (supplier == null) {
      System.out.println("Bookmark supplier could not be created");
      return null;
    }

    XNameAccess namedBookmarks = supplier.getBookmarks();
    if (namedBookmarks == null)  {
      System.out.println("Name access to bookmarks not possible");
      return null;
    }

    if (!namedBookmarks.hasElements())  {
      System.out.println("No bookmarks found");
      return null;
    }

    // find the specified bookmark
    Object oBookmark = null;
    try {
      oBookmark = namedBookmarks.getByName(bmName);
    }
    catch(com.sun.star.uno.Exception e) {}

    if (oBookmark == null) {
      System.out.println("Bookmark \"" + bmName + "\" not found");
      return null;
    }

    // there's no XBookmark, so return XTextContent
    return Lo.qi(XTextContent.class, oBookmark);
  }  // end of findBookmark()




  public static void addTextFrame(XTextCursor cursor, int yPos,
                                      String text, int width, int height)
  {
    try {
      XTextFrame xFrame = Lo.createInstanceMSF(XTextFrame.class, 
                                                 "com.sun.star.text.TextFrame");
      if (xFrame == null) {
        System.out.println("Could not create a text frame");
        return;
      }

      XShape tfShape = Lo.qi(XShape.class, xFrame);

      // set dimensions of the text frame
      tfShape.setSize(new Size(width, height));

      // anchor the text frame
      XPropertySet frameProps = Lo.qi(XPropertySet.class, xFrame);
      frameProps.setPropertyValue("AnchorType", // TextContentAnchorType.AS_CHARACTER);
                                              TextContentAnchorType.AT_PAGE);
      frameProps.setPropertyValue("FrameIsAutomaticHeight", true);   // will grow if necessary

      // add a red border around all 4 sides
      BorderLine border = new BorderLine();
      border.OuterLineWidth = 1;
      border.Color = 0xFF0000;  // red
            
      frameProps.setPropertyValue("TopBorder", border);
      frameProps.setPropertyValue("BottomBorder", border);
      frameProps.setPropertyValue("LeftBorder", border);
      frameProps.setPropertyValue("RightBorder", border);

      // make the text frame blue
      frameProps.setPropertyValue("BackTransparent", false);  // not transparent
      frameProps.setPropertyValue("BackColor", 0xCCCCFF);   // light blue

      // Set the horizontal and vertical position
      frameProps.setPropertyValue("HoriOrient", HoriOrientation.RIGHT);
                                             // HoriOrientation.NONE);
      // frameProps.setPropertyValue("HoriOrientPosition", 5000);

      frameProps.setPropertyValue("VertOrient", VertOrientation.NONE);
                                             // VertOrientation.CENTER);
      frameProps.setPropertyValue("VertOrientPosition", yPos);   // down from top


      // insert text frame into document (order is important here)
      append(cursor, xFrame);
      endParagraph(cursor);

      // add text into the text frame
      XText xFrameText = xFrame.getText();
      XTextCursor xFrameCursor = xFrameText.createTextCursor();
      xFrameText.insertString(xFrameCursor, text, false);
    }
    catch (Exception e) {
      System.out.println("Insertion of text frame failed: " + e);
    }
  }  // end of addTextFrame()





  public static void addTable(XTextCursor cursor, ArrayList<String[]> rowsList)
  /*  Each row becomes a row of the table. The first row is treated as a header,
      and colored in dark blue, and the rest in light blue. 
      Only 26 rows are allowed to keep the cell naming easy.
  */
  {
    try {
      // create a text table
      XTextTable table = Lo.createInstanceMSF(XTextTable.class, 
                                                 "com.sun.star.text.TextTable");
      if (table == null) {
        System.out.println("Could not create a text table");
        return;
      }

      // initialize the table dimensions
      int numRows = rowsList.size();
      int numCols = (rowsList.get(0)).length;
      if (numCols > 26) {    // column labelling goes from 'A' to 'Z'
        System.out.println("Too many columns: " + numCols + "; using first 26");
        numCols = 26;
      }
      System.out.println("Creating table rows: " + numRows + ", cols: " + numCols);
      table.initialize(numRows, numCols);

      // insert the table into the document
      append(cursor, table);
      endParagraph(cursor);

      // set table properties
      XPropertySet tableProps = Lo.qi(XPropertySet.class, table);
      tableProps.setPropertyValue("BackTransparent", false);  // not transparent
      tableProps.setPropertyValue("BackColor", 0xCCCCFF);   // light blue

      // set color of first row (i.e. the header) to be dark blue
      XTableRows rows = table.getRows();
      Props.setProperty(rows.getByIndex(0), "BackColor", 0x666694);     // dark blue

      // write table header
      String[] rowData = rowsList.get(0);
      for (int x=0; x < numCols; x++)
        setCellHeader( mkCellName(x,1), rowData[x], table);
                          // e.g. "A1", "B1", "C1", etc

      // insert table body
      for (int y=1; y < numRows; y++) {   // start in 2nd row
        rowData = rowsList.get(y);
        for (int x=0; x < numCols; x++)
          setCellText( mkCellName(x,y+1), rowData[x], table);
                           // e.g. "A2", "B5", "C3", etc
      }
    }
    catch (Exception e)
    {  System.out.println("Table insertion failed:" + e);  }
  }  // end of addTable()


  public static String mkCellName(int x, int y)
  // converts (x,y) to string (A+x) followed by y 
  {  return "" + ((char)('A' + x)) + y;  }



  private static void setCellHeader(String cellName, String data, XTextTable table)
  // puts text into the named cell of the table, colored white
  {
    XText cellText = Lo.qi(XText.class, table.getCellByName(cellName));
    XTextCursor textCursor = cellText.createTextCursor();
    Props.setProperty(textCursor, "CharColor", 0xFFFFFF);  // use white text

    cellText.setString(data);
  }  // end of setCellHeader()



  private static void setCellText(String cellName, String data, XTextTable table)
  // puts text into the named cell of the table
  {
    XText cellText = Lo.qi(XText.class, table.getCellByName(cellName));
    cellText.setString(data);
  }  // end of setCellText()



  // ================ add image in two ways ==========


  public static void addImageLink(XTextDocument doc, XTextCursor cursor, String fnm) 
  {  addImageLink(doc, cursor, fnm, 0, 0);  }


  public static void addImageLink(XTextDocument doc, XTextCursor cursor, 
  //                                               String fnm, double scaleFactor) 
                                               String fnm, int width, int height)
  // 1) image is inserted as a link
  {
    try {
      // create TextContent for graphic
      XTextContent tgo = Lo.createInstanceMSF(XTextContent.class, 
                                              "com.sun.star.text.TextGraphicObject");
      if (tgo == null) {
        System.out.println("Could not create a text graphic object");
        return;
      }

      // set anchor and link to file
      XPropertySet props =  Lo.qi( XPropertySet.class, tgo);
      props.setPropertyValue("AnchorType", TextContentAnchorType.AS_CHARACTER);
      props.setPropertyValue("GraphicURL", FileIO.fnmToURL(fnm));
            
      // optionally set the width and height
      if ((width > 0) && (height > 0)) {
        props.setPropertyValue("Width", width);
        props.setPropertyValue("Height", height);
      }

      // append image to document, followed by a newline
      append(cursor, tgo);
      endLine(cursor);
    }
    catch (Exception e) {
      System.out.println("Insertion of graphic in \"" + fnm + "\" failed: " + e);
    }
  }  // end of addImageLink()




  public static void addImageShape(XTextDocument doc, XTextCursor cursor, String fnm)
  {  addImageShape(doc, cursor, fnm, 0, 0);  }



  public static void addImageShape(XTextDocument doc, XTextCursor cursor,
                                                 String fnm, int width, int height)
  // 2) add image as a bitmap container inside a shape 
  {
    Size imSize;
    if ((width > 0) && (height > 0))
      imSize = new Size(width, height);
    else {
      imSize = Images.getSize100mm(fnm);   // in 1/100 mm units
      if (imSize == null)
        return;
    }

    try {
      // create TextContent for an empty graphic
      XTextContent gos = Lo.createInstanceMSF(XTextContent.class, 
                                       "com.sun.star.drawing.GraphicObjectShape");
      if (gos == null) {
        System.out.println("Could not create a graphic object shape");
        return;
      }

      // store the image's bitmap as the graphic shape's URL's value
      String bitmap = Images.getBitmap(fnm);
      Props.setProperty(gos, "GraphicURL", bitmap);

      // set the shape's size
      XShape xDrawShape = Lo.qi(XShape.class, gos);
      xDrawShape.setSize(imSize);   // must be included, or image is miniscule

      // insert image shape into the document, followed by newline
      append(cursor, gos);
      endLine(cursor);
    }
    catch(Exception e) {
      System.out.println("Insertion of graphic in \"" + fnm + "\" failed: " + e);
    }
  }  // end of addImageShape()




  public static void addLineDivider(XTextCursor cursor, int lineWidth)
  {
    try {
      // create TextContent for a line
      XTextContent ls = Lo.createInstanceMSF(XTextContent.class, 
                                                   "com.sun.star.drawing.LineShape");
      if (ls == null) {
        System.out.println("Could not create a line shape");
        return;
      }

      XShape lineShape = Lo.qi(XShape.class, ls);

      // lineShape.setPosition( new Point(0,0)); 
      lineShape.setSize( new Size(lineWidth, 0));    // units = 1/100 mm
      
      endParagraph(cursor);
      append(cursor, ls);
      endParagraph(cursor);

      // center the previous paragraph
      stylePrevParagraph(cursor, "ParaAdjust", 
                                  com.sun.star.style.ParagraphAdjust.CENTER);

      endParagraph(cursor);
    }
    catch(Exception e)
    {  System.out.println("Insertion of graphic line failed");  }

  }  // end of addLineDivider()





  // =================== extracting graphics from text doc ================

/*
  public static ArrayList<BufferedImage> getImages(XTextDocument textDoc)
  // return XGraphic images as BufferedImages;
  // requires temporary files (a bit ugly)
  {
    ArrayList<XGraphic> pics = getTextGraphics(textDoc);
    if ((pics == null) || (pics.size() == 0)) {
      System.out.println("No image found in document");
      return null;
    }

    ArrayList<BufferedImage> ims = new ArrayList<BufferedImage>();
    String tempFnm;
    for(XGraphic pic : pics) {
      tempFnm = FileIO.createTempFile("png");
      if (tempFnm != null) {
        Images.saveGraphic(pic, tempFnm, "png");
        BufferedImage im = Images.loadImage(tempFnm);
        FileIO.deleteFile(tempFnm);
        if (im != null)
          ims.add(im);
      }
    }
    return ims;
  }  // end of getImages()
*/



  public static ArrayList<XGraphic> getTextGraphics(XTextDocument textDoc)
  {
    XNameAccess xNameAccess = getGraphicLinks(textDoc);
    if (xNameAccess == null)
      return null;
    String[] names = xNameAccess.getElementNames();
    // System.out.println("Number of graphics found: " + names.length);

    ArrayList<XGraphic> pics = new ArrayList<XGraphic>();
    for (int i = 0; i < names.length; i++) {
      Object graphicLink = null;
      try {
        graphicLink = xNameAccess.getByName(names[i]);
      }
      catch(com.sun.star.uno.Exception e) {}

      if (graphicLink == null)
        System.out.println("No graphic found for " + names[i]);
      else {
        XGraphic xGraphic = Images.loadGraphicLink(graphicLink);
        if (xGraphic != null)
          pics.add(xGraphic);
        else
          System.out.println(names[i] + " could not be accessed");
      }
    }
    return pics;
  }  // end of getTextGraphics()



  public static XNameAccess getGraphicLinks(XComponent doc)
  {
    XTextGraphicObjectsSupplier imsSupplier = 
                Lo.qi(XTextGraphicObjectsSupplier.class, doc);
    if (imsSupplier == null) {
      System.out.println("Text graphics supplier could not be created");
      return null;
    }

    XNameAccess xNameAccess = imsSupplier.getGraphicObjects();
    if (xNameAccess == null)  {
      System.out.println("Name access to graphics not possible");
      return null;
    }

    if (!xNameAccess.hasElements())  {
      System.out.println("No graphics elements found");
      return null;
    }

    return xNameAccess;
  }  // end of getGraphicLinks()



  public static boolean isAnchoredGraphic(Object oGraphic)
  // look for graphic anchored in the document
  {
    XServiceInfo servInfo = Lo.qi(XServiceInfo.class, oGraphic);
    return (servInfo != null && 
        servInfo.supportsService("com.sun.star.text.TextContent") &&
        servInfo.supportsService("com.sun.star.text.TextGraphicObject"));
  }  // end of isAnchoredGraphic()




  public static XDrawPage getShapes(XTextDocument textDoc)
  {
    XDrawPageSupplier drawPageSupplier = Lo.qi(XDrawPageSupplier.class, textDoc);
    if (drawPageSupplier == null) {
      System.out.println("Draw page supplier could not be created");
      return null;
    }

    // get the draw page
    return drawPageSupplier.getDrawPage();

/*
    XTextShapesSupplier shpsSupplier = Lo.qi(XTextShapesSupplier.class, textDoc);
    if (shpsSupplier == null) {
      System.out.println("Text shapes supplier could not be created");
      return null;
    }

    XIndexAccess inAccess = shpsSupplier.getShapes();
    if (inAccess == null)  {
      System.out.println("Index access to shapes not possible");
      return null;
    }

    System.out.println("Num. shapes found: " + inAccess.getCount());
    return inAccess;
*/
  }  // end of getShapes()



  // -----------------  Linguistic API ------------------------


  public static void printServicesInfo(XLinguServiceManager2 lingoMgr)
  {
    com.sun.star.lang.Locale loc = 
                    new com.sun.star.lang.Locale("en","US","");

    System.out.println("Available Services:");
    printAvailServiceInfo(lingoMgr, "SpellChecker", loc);
    printAvailServiceInfo(lingoMgr, "Thesaurus", loc);
    printAvailServiceInfo(lingoMgr, "Hyphenator", loc);
    printAvailServiceInfo(lingoMgr, "Proofreader", loc);
    System.out.println();

    System.out.println("Configured Services:");
    printConfigServiceInfo(lingoMgr, "SpellChecker", loc);
    printConfigServiceInfo(lingoMgr, "Thesaurus", loc);
    printConfigServiceInfo(lingoMgr, "Hyphenator", loc);
    printConfigServiceInfo(lingoMgr, "Proofreader", loc);
    System.out.println();

    printLocales("SpellChecker", lingoMgr.getAvailableLocales("com.sun.star.linguistic2.SpellChecker"));
    printLocales("Thesaurus", lingoMgr.getAvailableLocales("com.sun.star.linguistic2.Thesaurus"));
    printLocales("Hyphenator", lingoMgr.getAvailableLocales("com.sun.star.linguistic2.Hyphenator"));
    printLocales("Proofreader", lingoMgr.getAvailableLocales("com.sun.star.linguistic2.Proofreader"));
    System.out.println();

  }  // end of printServicesInfo()



  public static void printAvailServiceInfo(XLinguServiceManager2 lingoMgr, String service, 
                                                          com.sun.star.lang.Locale loc)
  {
    String[] serviceNames = lingoMgr.getAvailableServices("com.sun.star.linguistic2." + service, loc);
    System.out.println(service + " (" + serviceNames.length + "):"); 
    for (String serviceName : serviceNames)
      System.out.println("  " + serviceName);
  }  // end of printAvailServiceInfo()



  public static void printConfigServiceInfo(XLinguServiceManager2 lingoMgr, String service, 
                                                           com.sun.star.lang.Locale loc)
  {
    String[] serviceNames = lingoMgr.getConfiguredServices("com.sun.star.linguistic2." + service, loc);
    System.out.println(service + " (" + serviceNames.length + "):"); 
    for (String serviceName : serviceNames)
      System.out.println("  " + serviceName);
  }  // end of printConfigServiceInfo()



  public static void printLocales(String service, com.sun.star.lang.Locale[] locs)
  {
    String[] countries = new String[locs.length];
    for(int i=0; i < locs.length; i++)
      countries[i] = locs[i].Country;
    Arrays.sort(countries);

    System.out.print("Locales for " + service + " (" + locs.length + "): ");
    for(int i=0; i < countries.length; i++) {
      if (i%10 == 0)
        System.out.println();
      System.out.print("  " + countries[i]);
    }
    System.out.println();
    System.out.println();
  }  // end of printLocales()



  public static void setConfiguredServices(XLinguServiceManager2 lingoMgr,
                                                 String service, String implName)
  { com.sun.star.lang.Locale loc = 
                    new com.sun.star.lang.Locale("en","US","");
    String[] implNames = { implName };
    lingoMgr.setConfiguredServices("com.sun.star.linguistic2." + service, loc, implNames);
  }  // end of setConfiguredServices()



  public static void dictsInfo()
  {
    // dictionaries are usually in <OFFICE>\share\wordbook
    XSearchableDictionaryList dictList = 
             Lo.createInstanceMCF(XSearchableDictionaryList.class, 
                     "com.sun.star.linguistic2.DictionaryList");
    if (dictList == null)
      System.out.println("No list of dictionaries found");
    else
      printDictsInfo(dictList);

    XConversionDictionaryList cdList = 
             Lo.createInstanceMCF(XConversionDictionaryList.class, 
                     "com.sun.star.linguistic2.ConversionDictionaryList");
    if (cdList == null)
      System.out.println("No list of conversion dictionaries found");
    else
      printConDictsInfo(cdList);
  }  // end of dictsInfo()



  public static void printDictsInfo(XSearchableDictionaryList dictList)
  {
    if (dictList == null) {
      System.out.println("Dictionary list is null");
      return;
    }

    System.out.println("No. of dictionaries: " + dictList.getCount());
    XDictionary[] dicts = dictList.getDictionaries();
    for(XDictionary dict : dicts)
      System.out.println("  " + dict.getName() + " (" + dict.getCount() + 
                                "); " + (dict.isActive() ? "active" : "na") +
                                "; \"" + dict.getLocale().Country + 
                                "\"; " + getDictType(dict.getDictionaryType()));
    System.out.println();
  }  // end of printDictsInfo()



  public static String getDictType(DictionaryType dt)
  {
    if (dt == DictionaryType.POSITIVE)
      return "positive";
    else if (dt == DictionaryType.NEGATIVE)
      return "negative";
    else if (dt == DictionaryType.MIXED)
      return "mixed";
    else
      return "??";
  }  // end of getDictType()



  public static void printConDictsInfo(XConversionDictionaryList cdList)
  {
    if (cdList == null) {
      System.out.println("Conversion Dictionary list is null");
      return;
    }

    XNameContainer dcCon = cdList.getDictionaryContainer();	
    String[] dcNames = dcCon.getElementNames();
    System.out.println("No. of conversion dictionaries: " + dcNames.length);
    for(String dcName : dcNames)
      System.out.println("  " + dcName);
    System.out.println();
  }  // end of printConDictsInfo()



  public static XLinguProperties getLinguProperties()
  {   return Lo.createInstanceMCF(XLinguProperties.class, 
                     "com.sun.star.linguistic2.LinguProperties");  }


  // ---------------- Linguistics: spell checking --------------

  public static XSpellChecker loadSpellChecker()
  {
    XLinguServiceManager lingoMgr = 
             Lo.createInstanceMCF(XLinguServiceManager.class, 
                     "com.sun.star.linguistic2.LinguServiceManager");
    if (lingoMgr == null) {
      System.out.println("No linguistics manager found");
      return null;
    }
    else
      return lingoMgr.getSpellChecker();
  }  // end of loadSpellChecker()



  public static int spellSentence(String sent, XSpellChecker speller)
  {
    String[] words = sent.split("\\W+");   // split into words
    int count = 0;
    boolean isCorrect;
    for(String word : words) {
      isCorrect = spellWord(word, speller);
      count = count + (isCorrect? 0 : 1);
    }
    return count;
  }  // end of spellSentence()



  public static boolean spellWord(String word, XSpellChecker speller)
  {
    com.sun.star.lang.Locale loc = 
             new com.sun.star.lang.Locale("en", "US", "");  // American English
    PropertyValue[] props = new PropertyValue[0];
    // PropertyValue[] props = Props.makeProps("IsSpellCapitalization", false);  // no effect

    XSpellAlternatives alts = speller.spell(word, loc, props);
    if (alts != null) {
      System.out.println("* \"" + word + "\" is unknown. Try:");
      String[] altWords = alts.getAlternatives();
      Lo.printNames(altWords);
      return false;
    }
    else
      return true;
  }   // end of spellWord()


  // ---------------- Linguistics: thesaurus --------------

  public static XThesaurus loadThesaurus()
  {
    XLinguServiceManager lingoMgr = 
             Lo.createInstanceMCF(XLinguServiceManager.class, 
                     "com.sun.star.linguistic2.LinguServiceManager");
    if (lingoMgr == null) {
      System.out.println("No linguistics manager found");
      return null;
    }
    else
      return lingoMgr.getThesaurus();
  }  // end of loadThesaurus()



  public static int printMeaning(String word, XThesaurus thesaurus)
  {
    com.sun.star.lang.Locale loc = 
                 new com.sun.star.lang.Locale("en", "US", "");  // American English
    PropertyValue[] props = new PropertyValue[0];

    XMeaning[] meanings = thesaurus.queryMeanings(word, loc, props);
    if (meanings == null) {
      System.out.println("\"" + word + "\" NOT found in thesaurus\n");
      return 0;
    }
    else {
      System.out.println("\"" + word + "\" found in thesaurus; number of meanings: " +
                                                meanings.length);
      for (int i=0; i < meanings.length; i++) {
        System.out.println((i+1) + ". Meaning: " + meanings[i].getMeaning());
        String[] synonyms = meanings[i].querySynonyms();
        System.out.println("  No. of synonyms: " + synonyms.length);
        for (int k=0; k < synonyms.length; k++)
          System.out.println("    " + synonyms[k]);
        System.out.println();
      }
      return meanings.length;
    }
  }  // end of printMeaning()



  // ---------------- Linguistics: grammar checking --------------

  public static XProofreader loadProofreader()
  {  return Lo.createInstanceMCF(XProofreader.class,
                                   "com.sun.star.linguistic2.Proofreader");  }



  public static int proofSentence(String sent, XProofreader proofreader)
  {
    com.sun.star.lang.Locale loc = 
             new com.sun.star.lang.Locale("en", "US", "");  // American English  
    PropertyValue[] props = new PropertyValue[0]; 
    int numErrs = 0;
    ProofreadingResult prRes = 
               proofreader.doProofreading("1", sent, loc, 0, sent.length(), props); 
    if (prRes != null) {
      SingleProofreadingError[] errs = prRes.aErrors;
      if (errs.length > 0)
        for(SingleProofreadingError err : errs) {
          printProofError(sent, err);
          numErrs++;
        }
    }
    return numErrs;
  }  // end of proofSentence()



  public static void printProofError(String str, SingleProofreadingError err)
  {
    String errText = str.substring(err.nErrorStart, err.nErrorStart+err.nErrorLength);
    System.out.println("G* " + err.aShortComment + " in: \"" + errText + "\"");
    if (err.aSuggestions.length > 0)
      System.out.println("   Suggested change: \"" + err.aSuggestions[0] + "\"");
    System.out.println();
  }  // end of printProofError()



  // ---------------- Linguistics: location guessing --------------

  public static com.sun.star.lang.Locale guessLocale(String testStr)
  {
    XLanguageGuessing guesser = Lo.createInstanceMCF(XLanguageGuessing.class,
                                     "com.sun.star.linguistic2.LanguageGuessing");
    if (guesser == null) {
      System.out.println("No language guesser found");
      return null;
    }
    else
      return guesser.guessPrimaryLanguage(testStr, 0, testStr.length());
  }  // end of guessLocale()



  public static void printLocale(com.sun.star.lang.Locale loc)
  // uses the Java Locale class since I couldn't get XLocale to work!
  {
    if (loc != null)  {
      //System.out.println("Locale lang: \"" + loc.Language + "\"; country: \"" + loc.Country +
      //                                  "\"; variant: \"" + loc.Variant + "\"");

      java.util.Locale jloc = new java.util.Locale(loc.Language, loc.Country, loc.Variant);
      System.out.println("Locale lang: \"" + jloc.getDisplayLanguage() + 
                         "\"; country: \"" + jloc.getDisplayCountry() +
                         "\"; variant: \"" + jloc.getDisplayVariant() + "\"");
    }
  }  // end of printLocale()



  // ---------------- Linguistics dialogs and menu items --------------


  public static void openSentCheckOptions()
  // open "Options - Language Settings - English sentence checking
  {
    XPackageInformationProvider pip = Info.getPip();
    String langExt = pip.getPackageLocation(
                            "org.openoffice.en.hunspell.dictionaries");
    System.out.println("Lang Ext: " + langExt);
    String url = langExt + "/dialog/en.xdl";

    PropertyValue[] props = Props.makeProps("OptionsPageURL", url);
    Lo.dispatchCmd("OptionsTreeDialog", props);
    Lo.wait(2000);
  }  // end of openSentCheckOptions()


  public static void openSpellGrammarDialog()
  // activate dialog in  Tools > Speling and Grammar...
  {  Lo.dispatchCmd("SpellingAndGrammarDialog");  
     Lo.wait(2000);    // is slow to load the first time
  }


  public static void toggleAutoSpellCheck()
  // toggle  Tools > Automatic Spell Checking
  {  Lo.dispatchCmd("SpellOnline");  }


  public static void openThesaurusDialog()
  // activate dialog in  Tools > Thesaurus...
  {  Lo.dispatchCmd("ThesaurusDialog");  }


}  // end of Write class



// Draw.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* A growing collection of utility functions to make Office
   easier to use. They are currently divided into the following
   groups:

     * open, create, save draw/impress doc
     * methods related to multiple slides/pages
     * layer management
     * view page management
     * master page methods
     * page methods
     * shape methods
     * draw/add shape to a page
     * get/set drawing properties 
     * draw an image
     * form manipulation
     * presentation related
*/

package utils;

import java.io.*;
// import java.awt.*;
import java.util.*;
import java.awt.image.*;
import javax.imageio.*;
import java.awt.geom.*;

import com.sun.star.beans.*;
import com.sun.star.comp.helper.*;
import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.text.*;
import com.sun.star.uno.*;
import com.sun.star.awt.*;
import com.sun.star.util.*;
import com.sun.star.drawing.*;
import com.sun.star.document.*;
import com.sun.star.container.*;
import com.sun.star.graphic.*;
import com.sun.star.sheet.*;
import com.sun.star.style.*;
import com.sun.star.table.*;
import com.sun.star.embed.*;
import com.sun.star.presentation.*;
import com.sun.star.animations.*;
import com.sun.star.form.*;
import com.sun.star.view.*;

import com.sun.star.uno.Exception;
import com.sun.star.io.IOException;


public class Draw
{

  //private static final String OFFICE_DIR = "C:/Program Files/LibreOffice 4/";

  private static final String SLIDE_TEMPLATE_PATH = "share/template/common/layout/";

  private static final int POLY_RADIUS = 20;

  //  index of the default glue points are 0 (top), 1 (right), 2 (bottom), and 3 (left)
  public static final int CONNECT_TOP = 0;
  public static final int CONNECT_RIGHT = 1;
  public static final int CONNECT_BOTTOM = 2;
  public static final int CONNECT_LEFT = 3;


  // slide layout names (see http://openoffice3.web.fc2.com/OOoBasic_Impress.html#OOoIPLy01a)
  // close to the names used in the Layouts panel in Impress
  public static final int LAYOUT_TITLE_SUB = 0;     // title, and subtitle below

  public static final int LAYOUT_TITLE_BULLETS = 1;   // the usual one you want

  public static final int LAYOUT_TITLE_CHART = 2;
  public static final int LAYOUT_TITLE_2CONTENT = 3;     
                     // 2 boxes: 1x2  (row x column), 1 row
  public static final int LAYOUT_TITLE_CONTENT_CHART = 4;
  public static final int LAYOUT_TITLE_CONTENT_CLIP = 6;
  public static final int LAYOUT_TITLE_CHART_CONTENT = 7;
  public static final int LAYOUT_TITLE_TABLE = 8;
  public static final int LAYOUT_TITLE_CLIP_CONTENT = 9;
  public static final int LAYOUT_TITLE_CONTENT_OBJECT = 10;

  public static final int LAYOUT_TITLE_OBJECT = 11;
  public static final int LAYOUT_TITLE_CONTENT_2CONTENT = 12;
                     // 3 boxes in 2 columns: 1 in first col, 2 in second
  public static final int LAYOUT_TITLE_OBJECT_CONTENT = 13;
  public static final int LAYOUT_TITLE_CONTENT_OVER_CONTENT = 14;   
                     // 2 boxes: 2x1, 1 column
  public static final int LAYOUT_TITLE_2CONTENT_CONTENT = 15;
                     // 3 boxes in 2 columns: 2 in first col, 1 in second
  public static final int LAYOUT_TITLE_2CONTENT_OVER_CONTENT = 16;   
                     // 3 boxes on 2 rows: 2 on first row, 1 on second
  public static final int LAYOUT_TITLE_CONTENT_OVER_OBJECT = 17;   
  public static final int LAYOUT_TITLE_4OBJECT = 18;      // 4 boxes: 2x2

  public static final int LAYOUT_TITLE_ONLY = 19;    // title only; no body shape
  public static final int LAYOUT_BLANK = 20;

  public static final int LAYOUT_VTITLE_VTEXT_CHART = 27;   
         // vertical title, vertical text, and chart
  public static final int LAYOUT_VTITLE_VTEXT = 28;   
  public static final int LAYOUT_TITLE_VTEXT = 29;   
  public static final int LAYOUT_TITLE_VTEXT_CLIP = 30;   

  public static final int LAYOUT_CENTERED_TEXT = 32;

  public static final int LAYOUT_TITLE_4CONTENT = 33;     // 4 boxes: 2x2
  public static final int LAYOUT_TITLE_6CONTENT = 34;     // 6 boxes: 2x3



  public static final String TITLE_TEXT = "com.sun.star.presentation.TitleTextShape";
  public static final String SUBTITLE_TEXT = "com.sun.star.presentation.SubtitleShape";
  public static final String BULLETS_TEXT = "com.sun.star.presentation.OutlinerShape";

  public static final String SHAPE_TYPE_NOTES = "com.sun.star.presentation.NotesShape";
  public static final String SHAPE_TYPE_PAGE = "com.sun.star.presentation.PageShape";

  // shape composition 
  public static final int MERGE = 0;
  public static final int INTERSECT = 1;
  public static final int SUBTRACT = 2;
  public static final int COMBINE = 3;


  // DrawPage slide show change constants
  public static final int CLICK_ALL_CHANGE = 0;
         // a mouse-click triggers the next animation effect or page change
  public static final int AUTO_CHANGE = 1;   
         // everything (page change, animation effects) is automatic
  public static final int CLICK_PAGE_CHANGE = 2;
        /* animation effects run automatically, but the user must click 
           on the page to change it */




  // ----------- open, create, save draw/impress doc ----------------------

  public static boolean isShapesBased(XComponent doc)
  {  return Info.isDocType(doc, Lo.DRAW_SERVICE) || 
            Info.isDocType(doc, Lo.IMPRESS_SERVICE);  }

  public static boolean isDraw(XComponent doc)
  {  return Info.isDocType(doc, Lo.DRAW_SERVICE);  }

  public static boolean isImpress(XComponent doc)
  {  return Info.isDocType(doc, Lo.IMPRESS_SERVICE);  }


  public static XComponent createDrawDoc(XComponentLoader loader)
  {  return Lo.createDoc("sdraw", loader);  }

  public static XComponent createImpressDoc(XComponentLoader loader)
  {  return Lo.createDoc("simpress", loader);  }

  public static String getSlideTemplatePath()
  {  return (Info.getOfficeDir() + SLIDE_TEMPLATE_PATH);  }



  public static void savePage(XDrawPage page, String fnm, String mimeType)
  // save the page in the specified file using the supplied office mime type
  {
    String saveFileURL = FileIO.fnmToURL(fnm);
    if (saveFileURL == null)
      return;
    System.out.println("Saving page in " + fnm);


    // create graphics exporter
    XGraphicExportFilter gef = Lo.createInstanceMCF(XGraphicExportFilter.class,
                       "com.sun.star.drawing.GraphicExportFilter");

    // set the output 'document' to be specified page
    XComponent doc = Lo.qi(XComponent.class, page);
    gef.setSourceDocument(doc);    // link exporter to the document

    // export the page by converting to the specified mime type
    PropertyValue props[] = Props.makeProps("MediaType", mimeType,
                                            "URL", saveFileURL);
    gef.filter(props);
    System.out.println("Export completed");
  }  // end of savePage()



// =========== methods related to document/multiple slides/pages ===================



  public static int getSlidesCount(XComponent doc)
  // get number of slide pages in the document
  {
    XDrawPages slides = getSlides(doc);
    if (slides == null)
      return 0;
    return slides.getCount();
  }  // end of getSlidesCount()




  public static XDrawPages getSlides(XComponent doc)
  // return all the slide pages as an XDrawPages sequence
  {
    XDrawPagesSupplier supplier = Lo.qi(XDrawPagesSupplier.class, doc);
    if (supplier == null)
      return null;
    return supplier.getDrawPages();
  }  // end of getSlides()



  public static ArrayList<XDrawPage> getSlidesList(XComponent doc)
  // return all the slide pages as an XDrawPage list
  {
    XDrawPages slides = getSlides(doc);
    if (slides == null)
      return null;
    int numSlides = slides.getCount();
    ArrayList<XDrawPage> slidesList = new ArrayList<XDrawPage>();
    try {
      for(int i=0; i < numSlides; i++)
        slidesList.add( Lo.qi(XDrawPage.class, slides.getByIndex(i)) );
    }
    catch(Exception e)
    {  System.out.println("Could not build slides array");  }
    return slidesList;
  }  // end of getSlidesList()



  public static XDrawPage[] getSlidesArr(XComponent doc)
  // return all the slide pages as an XDrawPage array
  {
    ArrayList<XDrawPage> slidesList = getSlidesList(doc);
    if (slidesList == null)
      return null;

    XDrawPage[] slidesArr = new XDrawPage[slidesList.size()];
    return slidesList.toArray(slidesArr);
  }  // end of getSlidesArr()



  public static XDrawPage getSlide(XComponent doc, int idx)
  {   return getSlide(getSlides(doc), idx);  }



  public static XDrawPage getSlide(XDrawPages slides, int idx)
  // get slide page by index
  {
    XDrawPage slide = null;
    try {
      slide = Lo.qi(XDrawPage.class, slides.getByIndex(idx));
    }
    catch(Exception e)
    {  System.out.println("Could not get slide " + idx);  }
    return slide;
  }  // end of getSlide()



  public int findSlideIdxByName(XComponent doc, String name)
  {
    int numSlides = getSlidesCount(doc);
    for (int i = 0; i < numSlides; i++) {
      XDrawPage slide = getSlide(doc, i);
      String nm = (String) Props.getProperty(slide, "LinkDisplayName");
      if (name.equalsIgnoreCase(nm))
        return i;
    }
    System.out.println("Could find slide " + name);
    return -1;
  }  // end of findSlideIdxByName()




  public static String getShapesText(XComponent doc)
  // return the text from inside all the document shapes
  {
    StringBuilder sb = new StringBuilder();
    ArrayList<XShape> xShapes = getOrderedShapes(doc);
    for(XShape xShape : xShapes) {
      String text = getShapeText(xShape);
      sb.append(text + "\n");
    }
    return sb.toString();
  }  // end of getShapesText()



  public static ArrayList<XShape> getShapes(XComponent doc)
  //  get all the shapes in all the pages of the doc
  {
    XDrawPage[] slides = getSlidesArr(doc);
    if (slides == null)
      return null;

    ArrayList<XShape> shapes = new ArrayList<XShape>(); 
    for(XDrawPage slide : slides)
      shapes.addAll( getShapes(slide));
    return shapes;
  }  // end of getShapes()



  public static ArrayList<XShape> getOrderedShapes(XComponent doc)
  //  get all the shapes in all the pages of the doc, in z-order per slide
  {
    XDrawPage[] slides = getSlidesArr(doc);
    if (slides == null)
      return null;

    ArrayList<XShape> shapes = new ArrayList<XShape>(); 
    for(XDrawPage slide : slides)
      shapes.addAll( getOrderedShapes(slide));
    return shapes;
  }  // end of getOrderedShapes()



  public static XDrawPage addSlide(XComponent doc)
  /* add a slide to the end of the document, and returns that page */
  {
    System.out.println("Adding a slide");
    XDrawPages slides = getSlides(doc);
    int numSlides = slides.getCount();
    return (XDrawPage) slides.insertNewByIndex(numSlides);
  }  // end of insertSlide()


  public static XDrawPage insertSlide(XComponent doc, int idx)
  /* inserts a slide at the given position in the document,
     and returns that page */
  {
    System.out.println("Inserting a slide at position " + idx);
    XDrawPages slides = getSlides(doc);
    return (XDrawPage) slides.insertNewByIndex(idx);
  }  // end of insertSlide()



  public static boolean deleteSlide(XComponent doc, int idx)
  /* inserts a blank slide page at the given position in the document,
     and returns that page */
  {
    System.out.println("Deleting a slide at position " + idx);
    XDrawPages slides = getSlides(doc);
    XDrawPage slide = null;
    try {
      slide = (XDrawPage) Lo.qi(XDrawPage.class, slides.getByIndex(idx));
    }
    catch(Exception e)
    {  System.out.println("Could not find slide " + idx);  
       return false;
    }

    slides.remove(slide);
    return true;
  }  // end of deleteSlide()



  public static XDrawPage duplicate(XComponent doc, int idx)
  {  
    XDrawPageDuplicator dup = Lo.qi(XDrawPageDuplicator.class, doc);  
    XDrawPage fromSlide = Draw.getSlide(doc, idx);
    if (fromSlide == null)
       return null;
    else
      return dup.duplicate(fromSlide);   // places copy after original
  }  // end of duplicate()


  // --------------------------- layer management ------------------------


  public static XLayerManager getLayerManager(XComponent doc)
  {
    XLayerSupplier xLayerSupplier = Lo.qi(XLayerSupplier.class, doc);
    XNameAccess xNameAccess = xLayerSupplier.getLayerManager();
    return Lo.qi(XLayerManager.class, xNameAccess);   
  }


  public static XLayer getLayer(XComponent doc, String layerName)
  {
    XLayerSupplier xLayerSupplier = Lo.qi(XLayerSupplier.class, doc);
    XNameAccess xNameAccess = xLayerSupplier.getLayerManager();
    try { 
      return Lo.qi(XLayer.class, xNameAccess.getByName(layerName));
    }
    catch(Exception e)
    {  System.out.println("Could not find the layer \"" + layerName + "\"");  
       return null;
    }
  }  // end of getLayer()



  public static XLayer addLayer(XLayerManager lm, String layerName)
  {
    XLayer layer = null;
    try {
      layer = lm.insertNewByIndex(lm.getCount());
      XPropertySet props = Lo.qi(XPropertySet.class, layer);
      props.setPropertyValue( "Name", layerName);
      props.setPropertyValue( "IsVisible", true);
      props.setPropertyValue( "IsLocked", false);
    }
    catch(Exception e)
    {  System.out.println("Could not add the layer \"" + layerName + "\"");  }
    return layer;
  }  // end of addLayer()



  // ------------------- view page ----------------------


  public static void gotoPage(XComponent doc, XDrawPage page) 
  { XController ctrl = GUI.getCurrentController(doc);
    gotoPage(ctrl, page);
  }  // end of jumpToPage()


  public static void gotoPage(XController ctrl, XDrawPage page) 
  { XDrawView xDrawView = Lo.qi(XDrawView.class, ctrl);
    xDrawView.setCurrentPage(page);
  }  // end of gotoPage()



  public static XDrawPage getViewedPage(XComponent doc)
  {
    XController ctrl = GUI.getCurrentController(doc);
    XDrawView xDrawView = Lo.qi(XDrawView.class, ctrl);
    return xDrawView.getCurrentPage();
  }  // end of getViewedPage()


	public static int getSlideNumber(XDrawView xDrawView) 
  {
		XDrawPage currPage = xDrawView.getCurrentPage();
    return Draw.getSlideNumber(currPage);
	}  // end of getSlideNumber()


  // ----------------------- master page methods ---------------


  public static int getMasterPageCount(XComponent doc)
  {
    XMasterPagesSupplier mpSupp = Lo.qi(XMasterPagesSupplier.class, doc);
    XDrawPages pgs = mpSupp.getMasterPages();
    return pgs.getCount();
  }


  public static XDrawPage getMasterPage(XComponent doc, int idx)
  // get master page by index
  {
    try {
      XMasterPagesSupplier mpSupp = Lo.qi(XMasterPagesSupplier.class, doc);
      XDrawPages pgs = mpSupp.getMasterPages();
      return Lo.qi(XDrawPage.class, pgs.getByIndex(idx));
    }
    catch(Exception e)
    {  System.out.println("Could not find master slide " + idx);  
       return null;
    }
  }  // end of getMasterPage()


  public static XDrawPage getMasterPage(XDrawPage slide)
  // return master page for the given slide
  {
    XMasterPageTarget mpTarget = Lo.qi(XMasterPageTarget.class, slide);
    return mpTarget.getMasterPage();
  }  // end of getMasterPage()


  public static XDrawPage insertMasterPage(XComponent doc, int idx)
  // creates new master page at the given index position,
  {
    XMasterPagesSupplier mpSupp = Lo.qi(XMasterPagesSupplier.class, doc);
    XDrawPages pgs = mpSupp.getMasterPages();
    return pgs.insertNewByIndex(idx);
  }  // end of insertMasterPage()


  public static void removeMasterPage(XComponent doc, XDrawPage slide)
  {
    XMasterPagesSupplier mpSupp = Lo.qi(XMasterPagesSupplier.class, doc);
    XDrawPages pgs = mpSupp.getMasterPages();
    pgs.remove(slide);
  }  // end of removeMasterPage()




  public static void setMasterPage(XDrawPage slide, XDrawPage mPg)
  // sets masterpage at the drawpage
  {
    XMasterPageTarget mpTarget = Lo.qi(XMasterPageTarget.class, slide);
    mpTarget.setMasterPage(mPg);
  }  // end of setMasterPage()



  public static XDrawPage getHandoutMasterPage(XComponent doc)
  {
    XHandoutMasterSupplier hmSupp = Lo.qi(XHandoutMasterSupplier.class, doc);
    return hmSupp.getHandoutMasterPage();
  }  // end of getHandoutMasterPage()



  public static XDrawPage findMasterPage(XComponent doc, String style)
  {
    try {
      XMasterPagesSupplier mpSupp = Lo.qi(XMasterPagesSupplier.class, doc);
      XDrawPages xMasterPages = mpSupp.getMasterPages();

      for (int i = 0; i < xMasterPages.getCount(); i++) {
        XDrawPage mPg = Lo.qi(XDrawPage.class, xMasterPages.getByIndex(i));
        String nm = (String) Props.getProperty(mPg, "LinkDisplayName");
        if (style.equals(nm))
          return mPg;
      }
      System.out.println("Could not find master slide " + style);  
      return null;
    }
    catch (Exception e) {
      System.out.println("Could not access master slide");  
      return null;
    }
  }  // end of findMasterPage()



  // ------------------ slide/page methods ----------------------------



  public static ArrayList<XShape> getShapes(XDrawPage slide)
  //  get all the shapes on the slide
  {
    if (slide == null) {
      System.out.println("Slide is null");
      return null;
    }
    if (slide.getCount() == 0) {
      System.out.println("Slide does not contain any shapes");
      return null;
    }

    ArrayList<XShape> shapes = new ArrayList<XShape>();
    try {
      for(int j=0; j < slide.getCount(); j++)
        shapes.add( Lo.qi(XShape.class, slide.getByIndex(j)));
    }
    catch(Exception e)
    {   System.out.println("Shapes extraction error in slide");  }

    return shapes;
  }  // end of getShapes()



  public static ArrayList<XShape> getOrderedShapes(XDrawPage slide)
  //  get all the shapes on the slide in increasing z-order
  {
    ArrayList<XShape> shapes = getShapes(slide);
    Collections.sort(shapes, new Comparator<XShape>() {
       public int compare(XShape s1, XShape s2) 
       {  return (getZOrder(s1) > getZOrder(s2)) ? -1 : 1;  }
    });

    return shapes;
  }  // end of getOrderedShapes()




  public static void showShapesInfo(XDrawPage slide)
  {
    System.out.println("Draw Page shapes:");
    ArrayList<XShape> shapes = getShapes(slide);
    if (shapes != null) {
      for(XShape shape : shapes)
        showShapeInfo(shape);
    }
  }  // end of showShapesInfo()



  public static String getShapeText(XDrawPage slide)
  // return all the text from inside the slide
  {
    StringBuilder sb = new StringBuilder();
    ArrayList<XShape> xShapes = getOrderedShapes(slide);
    for(XShape xShape : xShapes) {
      String text = getShapeText(xShape);
      sb.append(text + "\n");
    }
    return sb.toString();
  }  // end of getShapeText()



  public static int getSlideNumber(XDrawPage slide)
  {  return (Short) Props.getProperty(slide, "Number");  } 



  public static String getSlideTitle(XDrawPage slide)
  {
    XShape shape = findShapeByType(slide, TITLE_TEXT);
    if (shape == null)
      return null;
    else 
      return getShapeText(shape);
  }  // end of getSlideTitle()





  public static Size getSlideSize(XDrawPage xDrawPage)
  // get size of the given slide page (in mm units)
  {
    try {
      XPropertySet props = Lo.qi(XPropertySet.class, xDrawPage);
      if (props == null) {
         System.out.println("No slide properties found");
         return null;
       }
      int width = (Integer)props.getPropertyValue("Width");
      int height = (Integer)props.getPropertyValue("Height");
      return new Size(width/100, height/100);
    }
    catch(Exception e)
    {  System.out.println("Could not get page dimensions");
       return null;
    }
  }  // end of getSlideSize()



  public static void setName(XDrawPage currSlide, String name)
  {
    XNamed xPageName = Lo.qi(XNamed.class, currSlide);
    xPageName.setName(name);
  }  // end of setName()




  public static void titleSlide(XDrawPage currSlide, 
                                             String title, String subTitle)
  /* Add text to the slide page by treating it as a title page, which
     has two text shapes: one for the title, the other for a subtitle
  */
  {
    Props.setProperty(currSlide, "Layout", LAYOUT_TITLE_SUB);  
                                          // title, and subtitle below

    // XShapes xShapes = Lo.qi(XShapes.class, currSlide);

    // add the title text to the title shape
    // XShape xs = Lo.qi(XShape.class, xShapes.getByIndex(0));
    XShape xs = Draw.findShapeByType(currSlide, Draw.TITLE_TEXT);
    XText textField = Lo.qi(XText.class, xs);
    textField.setString(title);

    // add the subtitle text to the subtitle shape
    xs = Draw.findShapeByType(currSlide, Draw.SUBTITLE_TEXT);
    // xs = Lo.qi(XShape.class, xShapes.getByIndex(1));
    textField = Lo.qi(XText.class, xs);
    textField.setString(subTitle);
  }  // end of titleSlide()




  public static XText bulletsSlide(XDrawPage currSlide, String title)
  /* Add text to the slide page by treating it as a bullet page, which
     has two text shapes: one for the title, the other for a sequence of
     bullet points; add the title text but return a reference to the bullet
     text area
  */
  {
    Props.setProperty(currSlide, "Layout", LAYOUT_TITLE_BULLETS); 

    // add the title text to the title shape
    XShape xs = Draw.findShapeByType(currSlide, Draw.TITLE_TEXT);
    XText textField = Lo.qi(XText.class, xs);
    textField.setString(title);

    // return a reference to the bullet text area
    xs = Draw.findShapeByType(currSlide, Draw.BULLETS_TEXT);

    // print props info
    // Props.showObjProps("Outline Shape", xs);
    // Props.showIndexedProps("NumberingRules", Props.getProperty(xs, "NumberingRules"));

    return Lo.qi(XText.class, xs);
  }  // end of bulletsSlide()



  public static void addBullet(XText bullsText, int level, String text)
  /* add bullet text to the end of the bullets text area, specifying
     the nesting of the bullet using a numbering level value
     (numbering starts at 0).
  */
  {
    // access the end of the bullets text
    XTextRange bullsTextEnd = Lo.qi( XTextRange.class,bullsText).getEnd();
    // Props.showObjProps("TextRange in OutlinerShape", bullsTextEnd);

    // set the bullet's level
    Props.setProperty(bullsTextEnd, "NumberingLevel", (short)level);
    // Props.setProperty(bullsTextEnd, "NumberingIsNumber", false);

    bullsTextEnd.setString(text + "\n");  // add the text
  }  // end of addBullet()




  public static void titleOnlySlide(XDrawPage currSlide, String header)
  // create a slide with only a title
  {
    Props.setProperty(currSlide, "Layout", LAYOUT_TITLE_ONLY); 
                                         // title only; no body shape

    // add the text to the title shape
    XShape xs = Draw.findShapeByType(currSlide, Draw.TITLE_TEXT);
    XText textField = Lo.qi(XText.class, xs);
    textField.setString(header);
  }  // end of titleOnlySlide()



  public static void blankSlide(XDrawPage currSlide)
  {  Props.setProperty(currSlide, "Layout", LAYOUT_BLANK);  }  



/*
  public static void copyTo(XDrawPage fromSlide, XDrawPage toSlide)
  // All the shapes on the fromSlide are copied over to the toSlide
  {
    blankSlide(toSlide);
    ArrayList<XShape> shapes = getShapes(fromSlide);
    System.out.println("No. of shapes being copied: " + shapes.size());
    for(XShape shape : shapes)
      copyShape(toSlide, shape);
  }  // end of copyTo()
*/



  static public XDrawPage getNotesPage(XDrawPage slide)
  // each draw page has a notes page
  {
    XPresentationPage presPage = Lo.qi(XPresentationPage.class, slide);
    if (presPage == null) {
      System.out.println("This is not a presentation slide, so no notes page is available");
      return null;
    }
    return presPage.getNotesPage();
  }  // end of getNotesPage()



  public XDrawPage getNotesPageByIndex(XComponent doc, int index)
  {
    XDrawPage slide = getSlide(doc, index);
    return getNotesPage(slide);
  }




  // ======================== shape methods ====================



  public static void showShapeInfo(XShape xShape)
  {
    System.out.println("  Shape service: " + xShape.getShapeType() + 
                       "; z-order: " + getZOrder(xShape));
  }  // end of showShapeInfo()
    



  public static String getShapeText(XShape xShape)
  // get text from inside a shape
  {
    String text = null;
    XText xText = Lo.qi(XText.class, xShape);

    XTextCursor xTextCursor = xText.createTextCursor();
    XTextRange xTextRange = Lo.qi(XTextRange.class, xTextCursor);
    text = xTextRange.getString();
    return text;
  }  // end of getShapeText()



  public static XShape findShapeByType(XDrawPage slide, String shapeType)
  {
    ArrayList<XShape> shapes = getShapes(slide);
    if (shapes == null) {
      System.out.println("No shapes were found in the draw page");
      return null;
    }

    for (XShape shape : shapes) {
      if (shapeType.equals(shape.getShapeType()))
        return shape;
    }

    System.out.println("No shape found of type \"" + shapeType + "\"");
    return null;
  }  // end of findShapeByType()





  public static XShape findShapeByName(XDrawPage slide, String shapeName)
  {
    ArrayList<XShape> shapes = getShapes(slide);
    if (shapes == null) {
      System.out.println("No shapes were found in the draw page");
      return null;
    }

    for (XShape shape : shapes) {
      String nm = (String)Props.getProperty(shape, "Name");
      if (shapeName.equals(nm))
        return shape;
    }

    System.out.println("No shape named \"" + shapeName + "\"");
    return null;
  }  // end of findShapeByName()




  public static XShape copyShapeContents(XDrawPage slide, XShape oldShape)
  {
    XShape shape = copyShape(slide, oldShape);
    System.out.println("Shape type: " + oldShape.getShapeType());
  //  if (oldShape.getShapeType().equals("com.sun.star.drawing.TextShape"))
      addText(shape, getShapeText(oldShape));

    return shape;
  }  // end of copyShapeContents()



  public static XShape copyShape(XDrawPage slide, XShape oldShape)
  // parameters are in 1/100 mm units 
  { 
    Point pt = oldShape.getPosition();
    Size sz =  oldShape.getSize();

    XShape shape = null;
    try {
      shape = Lo.createInstanceMSF(XShape.class, oldShape.getShapeType());
      System.out.println("Copying: " + oldShape.getShapeType());
      // Props.setProperties(shape, oldShape);
      shape.setPosition(pt);
      shape.setSize(sz);
      slide.add(shape);
    }
    catch(Exception e)
    {  System.out.println("Unable to copy shape");  }

    return shape;
  }  // end of copyShape()



  public static void setZOrder(XShape shape, int order)
  {  Props.setProperty(shape, "ZOrder", order);  } 


  public static int getZOrder(XShape shape)
  {  return (Integer) Props.getProperty(shape, "ZOrder");  }


  public static void moveToTop(XDrawPage slide, XShape shape)
  { int maxZO = findBiggestZOrder(slide);
    setZOrder(shape, maxZO+1);
  } 



  public static int findBiggestZOrder(XDrawPage slide)
  {  return getZOrder( findTopShape(slide));  }



  public static XShape findTopShape(XDrawPage slide)
  {
    ArrayList<XShape> shapes = getShapes(slide);
    if ((shapes == null) || (shapes.size() == 0)) {
      System.out.println("No shapes found");
      return null;
    }
    int maxZOrder = 0;
    XShape sTop = null;
    for (XShape shape : shapes) {
      int zo = getZOrder(shape);
      if (zo > maxZOrder) {
        maxZOrder = zo;
        sTop = shape;
      }
    }
    return sTop;
  }  // end of findTopShape()



  public static void moveToBottom(XDrawPage slide, XShape shape)
  {
    ArrayList<XShape> shapes = getShapes(slide);
    if ((shapes == null) || (shapes.size() == 0)) {
      System.out.println("No shapes found");
      return;
    }

    int minZOrder = 999;
    for (XShape sh : shapes) {
      int zo = getZOrder(sh);
      if (zo < minZOrder)
        minZOrder = zo;
      setZOrder(sh, zo+1);
    }
    setZOrder(shape, minZOrder);
  }  // end of moveToBottom()





  // ================== draw/add shape to a page ================



  public static XShape drawRectangle(XDrawPage slide, int x, int y, int width, int height)
  {  return addShape(slide, "RectangleShape", x, y, width, height);  }



  public static XShape drawCircle(XDrawPage slide, int x, int y, int radius)
  {  return addShape(slide, "EllipseShape", x-radius, y-radius, radius*2, radius*2);  }



  public static XShape drawEllipse(XDrawPage slide, int x, int y, int width, int height)
  {  return addShape(slide, "EllipseShape", x, y, width, height);  }



  public static XShape drawPolygon(XDrawPage slide, int x, int y, int nSides)
  {  return drawPolygon(slide, x, y, POLY_RADIUS, nSides);  }


  public static XShape drawPolygon(XDrawPage slide, int x, int y, int radius, int nSides)
  {
    XShape polygon = addShape(slide, "PolyPolygonShape", 0, 0, 0, 0);
       // for shapes formed by one *or more* polygons
    
    Point[] pts = genPolygonPoints(x, y, radius, nSides);
    Point[][] polys = new Point[][] {pts};
       // could be many polygons pts in this 2D array
    Props.setProperty(polygon, "PolyPolygon", polys);
    return polygon;
  }  // end of drawPolygon()



  private static Point[] genPolygonPoints(int x, int y, int radius, int nSides)
  {
    if (nSides < 3) {
      System.out.println("Too few sides; must be 3 or more");
      nSides = 3;
    }
    else if (nSides > 30) {
      System.out.println("Too many sides; must be 30 or less");
      nSides = 30;
    }

    Point[] pts = new Point[nSides];
    double angleStep = Math.PI/nSides;
    for (int i = 0; i < nSides; i++) {
      pts[i] = new Point(
                 (int) Math.round(x*100 + radius*100*Math.cos(i*2*angleStep)),
                 (int) Math.round(y*100 + radius*100*Math.sin(i*2*angleStep)) );
    }
    return pts;
  }  // end of genPolygonPoints()



  public static XShape drawBezier(XDrawPage slide, 
                               Point[] pts, PolygonFlags[] flags, boolean isOpen)
  {
    if (pts.length != flags.length) {
      System.out.println("Mismatch in lengths of points and flags array");
      return null;
    }

    String bezierType = isOpen ? "OpenBezierShape" : "ClosedBezierShape";
    XShape bezierPoly = addShape(slide, bezierType, 0, 0, 0, 0);
    
    // create space for one bezier shape
    PolyPolygonBezierCoords aCoords = new PolyPolygonBezierCoords();
                // for shapes formed by one *or more* bezier polygons
    aCoords.Coordinates = new Point[1][];
    aCoords.Flags = new PolygonFlags[1][];
    aCoords.Coordinates[0] = pts;
    aCoords.Flags[0] = flags;

    Props.setProperty(bezierPoly, "PolyPolygonBezier", aCoords);
    return bezierPoly;
  }  // end of drawBezier()



  public static XShape drawLine(XDrawPage slide, int x1, int y1, int x2, int y2)
  {
    // System.out.println("Drawing a line");
    // make sure size is non-zero
    if ((x1 == x2) && (y1 == y2)) {
      System.out.println("Line is a point");
      return null;
    }

    int width = x2 - x1;   // may be negative
    int height = y2 - y1;  // may be negative
    // System.out.println("x1-y1: " + x1 + " - " + y1);
    // System.out.println("width-height: " + width + " - " + height);
    return addShape(slide, "LineShape",  x1, y1, width, height);
  }  // end of drawLine()



  public static XShape drawPolarLine(XDrawPage slide, int x, int y, int degrees, int distance)
  /* Draw a line from x,y in the direction of degrees, for the specified distance
     degrees is measured clockwise from x-axis
  */
  {
    int xDist = (int)Math.round(Math.cos(Math.toRadians(degrees)) * distance);
    int yDist = -(int)Math.round(Math.sin(Math.toRadians(degrees)) * distance);
    return drawLine(slide, x, y, x+xDist, y+yDist);
  }



  public static XShape drawLines(XDrawPage slide, int[] xs, int[] ys)
  {
    if (xs.length != ys.length) {
      System.out.println("The two arrays must be the same length");
      return null;
    }

    int numPoints = xs.length;
    Point[] pts = new Point[numPoints];
    for (int i = 0; i < numPoints; i++)
      pts[i] = new Point(xs[i]*100, ys[i]*100);   // so in 1/100 mm units

    Point[][] linePaths  = new Point[][] {pts};
       // an array of Point arrays, one Point array for each line path

    XShape polyLine = addShape(slide, "PolyLineShape", 0, 0, 0, 0);
       // for a shape formed by from multiple connected lines
   
    Props.setProperty(polyLine, "PolyPolygon", linePaths);
    return polyLine;
  }  // end of drawLines()




  public static XShape drawText(XDrawPage slide, String msg,
                    int x, int y, int width, int height)
  { XShape shape = addShape(slide, "TextShape",  x, y, width, height);
    addText(shape, msg, 0);
    return shape;
  }  // end of drawText()


  public static XShape drawText(XDrawPage slide, String msg,
                    int x, int y, int width, int height, int fontSize)
  { XShape shape = addShape(slide, "TextShape",  x, y, width, height);
    addText(shape, msg, fontSize);
    return shape;
  }  // end of drawText()



  public static void addText(XShape shape, String msg)
  {  addText(shape, msg, 0); }


  public static void addText(XShape shape, String msg, int fontSize)
  {
    XText xText = Lo.qi(XText.class, shape);
    XTextCursor cursor = xText.createTextCursor();
    cursor.gotoEnd(false);
    if (fontSize > 0)
      Props.setProperty(cursor, "CharHeight", fontSize);
    XTextRange range = Lo.qi(XTextRange.class, cursor);
    range.setString(msg);
  }  // end of addText()




  public static void addConnector(XDrawPage slide, XShape shape1, XShape shape2)
  {  addConnector(slide, shape1, CONNECT_RIGHT, shape2, CONNECT_LEFT);  }



  public static XShape addConnector(XDrawPage slide, 
                XShape shape1, int startConnect, XShape shape2, int endConnect)
  // connect indicies are 0 (top), 1 (right), 2 (bottom), and 3 (left)
  {
    XShape xConnector = addShape(slide, "ConnectorShape",  0, 0, 0, 0);

    XPropertySet props = Lo.qi(XPropertySet.class, xConnector);
    try { 
      props.setPropertyValue("StartShape", shape1);
      props.setPropertyValue("StartGluePointIndex", startConnect);
      
      props.setPropertyValue("EndShape", shape2);
      props.setPropertyValue("EndGluePointIndex", endConnect);

      props.setPropertyValue("EdgeKind", ConnectorType.STANDARD); 
                          // STANDARD, CURVE, LINE, LINES
    }
    catch(Exception e)
    {  System.out.println("Could not connect the shapes");  }

    return xConnector;
  }  // end of addConnectorShape()



  public static GluePoint2[] getGluePoints(XShape shape)
  {
    XGluePointsSupplier gpSupp = Lo.qi( XGluePointsSupplier.class, shape);
    XIndexContainer gluePts = gpSupp.getGluePoints();

    int numGPs = gluePts.getCount();  // should be 4 by default
    if (numGPs == 0) {
      System.out.println("No glue points for this shape");
      return null;
    }
    GluePoint2[] gps = new GluePoint2[numGPs];
    for(int i=0; i < numGPs; i++) {
      try {
        gps[i] = Lo.qi(GluePoint2.class, gluePts.getByIndex(i));
        // System.out.println("Glue point " + i + ": " + gp);
      }
      catch(com.sun.star.uno.Exception e)
      {  System.out.println("Could not access glue point " + i);  }
    }
    return gps;
  }  // end of getGluePoints()



  public static XShape getChartShape(XDrawPage slide,
                                       int x, int y, int width, int height)
  {  
    XShape shape = addShape(slide, "OLE2Shape", x, y, width, height);
    Props.setProperty(shape, "CLSID", Lo.CHART_CLSID);  // a chart
    return shape;
  }  // end of getChartShape()



  public static XShape drawFormula(XDrawPage slide, String formula,
                                       int x, int y, int width, int height)
  {  
    XShape shape = addShape(slide, "OLE2Shape", x, y, width, height);
    Props.setProperty(shape, "CLSID", Lo.MATH_CLSID);  // a formula
                                     
    XModel model = Lo.qi(XModel.class, Props.getProperty(shape, "Model") );
    // Info.showServices("OLE2Shape Model", model);
    Props.setProperty(model, "Formula", formula);   // from FormulaProperties
    return shape;
  }  // end of drawFormula()


  public static XShape drawMedia(XDrawPage slide, String fnm,
                                       int x, int y, int width, int height)
  // causes Office to crash on exiting
  {  
    XShape shape = addShape(slide, "MediaShape", x, y, width, height);

    Props.showObjProps("Shape", shape);
    System.out.println("Loading media: \"" + fnm + "\"");
    Props.setProperty(shape, "MediaURL", FileIO.fnmToURL(fnm)); 
    Props.setProperty(shape, "Loop", true); 

    // Props.setProperty(shape, "PlayFull", true); 
    // Props.setProperty(shape, "IsPlaceholderDependent", true); 

    return shape;
  }  // end of drawMedia()





  public static XShape addShape(XDrawPage slide, String shapeType, 
                                              int x, int y, int width, int height)
  { warnsPosition(slide, x, y);
    XShape shape = makeShape(shapeType, x, y, width, height);
    if (shape != null)
      slide.add(shape);
    return shape;
  }  // end of addShape()



  private static void warnsPosition(XDrawPage slide, int x, int y)
  // warns if (x, y) is not on the page
  {
    Size slideSize = Draw.getSlideSize(slide);
    if (slideSize == null) {
      System.out.println("No slide size found");
      return;
    }
    int slideWidth = slideSize.Width;
    int slideHeight = slideSize.Height;

    if (x < 0)
      System.out.println("x < 0");
    else if (x > slideWidth-1)
      System.out.println("x position off right hand side of the slide");

    if (y < 0)
      System.out.println("y < 0");
    else if (y > slideHeight-1)
      System.out.println("y position off bottom of the slide");
  }  // end of warnsPosition()



  public static XShape makeShape(String shapeType, int x, int y, int width, int height)
  // parameters are in mm units 
  { 
    XShape shape = null;
    try {
      shape = Lo.createInstanceMSF(XShape.class, "com.sun.star.drawing."+shapeType);
      shape.setPosition( new Point(x*100,y*100) );
      shape.setSize( new Size(width*100, height*100) );
    }
    catch(Exception e)
    {  System.out.println("Unable to create shape: " + shapeType);  }

    return shape;
  }  // end of makeShape()



  public static boolean isGroup(XShape shape)
  {  return (shape.getShapeType().equals("com.sun.star.drawing.GroupShape"));  } 



  public static XShape combineShape(XComponent doc, XShapes shapes, int combineOp)
  {
    // select the shapes for the dispatches to apply to
    XSelectionSupplier selSupp = Lo.qi(XSelectionSupplier.class, 
                                                GUI.getCurrentController(doc) );
    selSupp.select(shapes);

    if (combineOp == MERGE)
      Lo.dispatchCmd("Merge");
    else if (combineOp == INTERSECT)
      Lo.dispatchCmd("Intersect");
    else if (combineOp == SUBTRACT)
      Lo.dispatchCmd("Substract");   // misspelt!
    else if (combineOp == COMBINE)
      Lo.dispatchCmd("Combine"); 
    else {
       System.out.println("Did not recognize op: " + combineOp + "; using merge");
       Lo.dispatchCmd("Merge");
    }
    Lo.delay(500);   // give time for dispatches to arrive and be processed

    // extract the new single shape from the modified selection
    XShapes xs = Lo.qi(XShapes.class, selSupp.getSelection());
    XShape combinedShape = null;
    try {
      combinedShape = Lo.qi(XShape.class, xs.getByIndex(0));
      // System.out.println("Combined Shape type: " + combinedShape.getShapeType());
    }
    catch(com.sun.star.uno.Exception e)
    {  System.out.println("Could not get combined shape");  }

    return combinedShape;
  }  // end of combineShape()



  public static XControlShape createControlShape(String label, 
                           int x, int y, int width, int height, String shapeKind)
  {
    try {
      XControlShape cShape = Lo.createInstanceMSF(XControlShape.class, 
                                          "com.sun.star.drawing.ControlShape");
      cShape.setSize(new Size(width*100, height*100));
      cShape.setPosition(new Point(x*100, y*100));

      XControlModel cModel = Lo.createInstanceMSF(XControlModel.class, 
                                   "com.sun.star.form.component." + shapeKind);

      XPropertySet props = Lo.qi(XPropertySet.class, cModel);
      props.setPropertyValue("DefaultControl", "com.sun.star.form.control." + shapeKind);
      props.setPropertyValue("Name", "XXX");
      props.setPropertyValue("Label", label);
      // props.setPropertyValue("BackgroundColor", 0x444444);

      props.setPropertyValue("FontHeight", 18.0);   // used new Float()
      props.setPropertyValue("FontName", "Times");
/*
      FontDescriptor fd = (FontDescriptor) props.getPropertyValue("FontDescriptor");
      System.out.println("Font descriptor: " + fd.Name + "; " + fd.Height);
      float fontHeight = (Float) props.getPropertyValue("FontHeight");
      System.out.println("Font height: " + fontHeight);
*/

      cShape.setControl(cModel);

      // XControl xButtonControl = cShape.getControl();
      XButton xButton = Lo.qi(XButton.class, cModel);
      if (xButton == null)
        System.out.println("XButton is null");

      Props.showProps("Control model props", props);

      return cShape;
    }
    catch (Exception e) {
      System.out.println("Could not create control shape: " + e);
      return null;
    }
  } // end of createControlShape()



  // --------------- custom shape addition using dispatch and JNA --------


  public static XShape addDispatchShape(XDrawPage slide, String shapeDispatch, 
                                              int x, int y, int width, int height)
  // ((x,y), width, height must be set after insertion
  { warnsPosition(slide, x, y);
    XShape shape = createDispatchShape(slide, shapeDispatch);
    if (shape != null) {
      setPosition(shape, x, y);
      setSize(shape, width, height);
    }
    return shape;
  }  // end of addDispatchShape()




  public static XShape createDispatchShape(XDrawPage slide, String shapeDispatch)
  /*  Create a dispatch shape in two steps: select the shape by calling dispatchCmd()
      and then create it by imitating a press and drag on the visible page.

      A reference to the created shape is obtained by assuming that it's the new 
      top-most element on the page.

      shapeDispatch is the dispatch name for a shape (e.g. "BasicShapes.diamond").
      See dispatchShapes.txt
  */
  {
    int numShapes = slide.getCount();

    Lo.dispatchCmd(shapeDispatch);   // select the shape icon; Office must be visible
    Lo.wait(1000);

    // click and drag on the page to create the shape on the page;
    // the current page must be visible
    java.awt.Point p1 = JNAUtils.getClickPoint( JNAUtils.getHandle() );
    java.awt.Point p2 = JNAUtils.getOffsetPoint(p1, 100, 100);  // hardwired offset 
    JNAUtils.doDrag(p1, p2);  // drag the mouse between p1 and p2
    Lo.wait(2000);
 
    //Draw.showShapesInfo(slide);

    // get a reference to the shape by assuming it's the top one on the page
    int numShapes2 = slide.getCount();
    if (numShapes2 == numShapes+1) {
      System.out.println("Shape \"" + shapeDispatch + "\" created");
      return Draw.findTopShape(slide);
    }
    else {
      System.out.println("Shape \"" + shapeDispatch + "\" not created");
      return null;
    }
  }  // end of createDispatchShape()



 // ------------------------------ presentation shapes -----------------------


  public static void setMasterFooter(XDrawPage master, String text)
  // set the master page's footer text
  {
    XShape footerShape = Draw.findShapeByType(master, 
                                      "com.sun.star.presentation.FooterShape");
    XText textField = Lo.qi(XText.class, footerShape);
    textField.setString(text);
  }  // end of setMasterFooter()



  public static XShape addSlideNumber(XDrawPage slide)
  // add slide number at bottom right (like on the default slide)
  {
    Size sz = Draw.getSlideSize(slide);
    int width = 60;
    int height = 15;
    return Draw.addPresShape(slide, "SlideNumberShape", 
                          sz.Width-width-12, sz.Height-height-4, width, height); 
  }  // end of addSlideNumber()



  public static XShape addPresShape(XDrawPage slide, String shapeType, 
                                              int x, int y, int width, int height)
  // ((x,y), width, height must be set after insertion
  { warnsPosition(slide, x, y);
    XShape shape = Lo.createInstanceMSF(XShape.class, 
                                  "com.sun.star.presentation."+shapeType);
    if (shape != null) {
      slide.add(shape);
      setPosition(shape, x, y);
      setSize(shape, width, height);
    }
    return shape;
  }  // end of addPresShape()

    

  // ----------------------- get/set drawing properties --------------------------



  public static Point getPosition(XShape shape)
  { Point pt = shape.getPosition();
    return new Point(pt.X/100, pt.Y/100);  // convert to mm
  }  // end of getPosition()



  public static Size getSize(XShape shape)
  { Size sz = shape.getSize();
    return new Size(sz.Width/100, sz.Height/100);  // convert to mm
  }  // end of getSize()



  public static void printPoint(Point pt)
  // print the point in mm units
  {  System.out.println("  Point (mm): (" + pt.X/100 + ", " +  pt.Y/100 +  ")");  }


  public static void printSize(Size sz)
  // print the size in mm units
  {  System.out.println("  Size (mm): [" + sz.Width/100 + ", " +  sz.Height/100 +  "]");  }



  public static void reportPosSize(XShape shape)
  {
    if (shape == null) {
      System.out.println("The shape is null");
      return;
    }
    System.out.println("Shape name: " + Props.getProperty(shape, "Name"));
    System.out.println("  Type: " + shape.getShapeType());
    printPoint( shape.getPosition());
    printSize( shape.getSize());
  }  // end of reportPosSize()



  public static void setPosition(XShape shape, Point pt)
  {   shape.setPosition(new Point(pt.X*100, pt.Y*100));  }  // convert to 1/100 mm units

  public static void setPosition(XShape shape, int x, int y)
  {   shape.setPosition(new Point(x*100, y*100));  } 



  public static void setSize(XShape shape, Size sz)
  {  setSize(shape, sz.Width, sz.Height);  }


  public static void setSize(XShape shape, int width, int height)
  {
    try {
      shape.setSize(new Size(width*100, height*100));  // convert to 1/100 mm units
      System.out.println("(w,h): " + width + ", " + height);
    }
    catch(Exception e)
    { System.out.println("Could not set the shape's size");  }
  }  // end of setSize()



  public static void setStyle(XShape shape, XNameContainer graphicStyles,
                                                                String styleName)
  /* A list of graphics styles can be found at
      https://wiki.openoffice.org/wiki/Documentation/DevGuide/Drawings/Overall_Document_Features
  */
  { try { 
      XStyle style = Lo.qi(XStyle.class, graphicStyles.getByName(styleName));
      Props.setProperty(shape, "Style", style);
    }
    catch(Exception e)
    {  System.out.println("Could not set the style to \"" + styleName + "\"");  }
  }  // end setStyle()



  public static XPropertySet getTextProperties(XShape xShape)
  // return the properties associated with the text area inside the shape
  {
    XText xText = Lo.qi(XText.class, xShape);
    XTextCursor xTextCursor = xText.createTextCursor();
    xTextCursor.gotoStart(false);
    xTextCursor.gotoEnd(true);
    XTextRange xTextRange = Lo.qi(XTextRange.class, xTextCursor);
    return Lo.qi(XPropertySet.class, xTextRange);
  }  // end of getTextProperties()



  public static java.awt.Color getLineColor(XShape shape)
  {
    XPropertySet props = Lo.qi(XPropertySet.class, shape);
    try {
      int rgb = ((Integer)props.getPropertyValue("LineColor")).intValue();
      return new java.awt.Color(rgb);
    }
    catch(Exception e) {
      System.out.println("Could not access line color");
      return null;
    }
  }  // end of getLineColor()


  public static void setDashedLine(XShape shape, boolean isDashed)
  {
    LineDash ld = new LineDash();   // dashes only; no dots
    ld.Dots = 0; ld.DotLen = 100;
    ld.Dashes = 5;
    ld.DashLen = 200;
    ld.Distance = 200;

    XPropertySet props = Lo.qi(XPropertySet.class, shape);
    try {
      if (isDashed) {
        props.setPropertyValue("LineStyle", LineStyle.DASH);
        props.setPropertyValue("LineDash", ld);
      }
      else // switch to a solid line
        props.setPropertyValue("LineStyle", LineStyle.SOLID);
    }
    catch(Exception e)
    {  System.out.println("Could not set dashed line property");  }
  }  // end of setDashedLine()



  public static int getLineThickness(XShape shape)
  {
    XPropertySet props = Lo.qi(XPropertySet.class, shape);
    try {
      return ((Integer)props.getPropertyValue("LineWidth")).intValue();
    }
    catch(Exception e) {
      System.out.println("Could not access line thickness");
      return 0;
    }
  }  // end of getLineThickness()


  public static java.awt.Color getFillColor(XShape shape)
  {
    XPropertySet props = Lo.qi(XPropertySet.class, shape);
    try {
      int rgb = ((Integer)props.getPropertyValue("FillColor")).intValue();
      return new java.awt.Color(rgb);
    }
    catch(Exception e) {
      System.out.println("Could not access fill color");
      return null;
    }
  }  // end of getFillColor()


  public static void setTransparency(XShape shape, int level)
  // higher level means more transparent
  {
    if ((level < 0) || (level > 100)) {
      System.out.println("Transparency level must be between 0-100; using 50");
      level = 50;
    }
    Props.setProperty(shape, "FillTransparence", level);
  }  // end of setTransparency()



  public static void setGradientColor(XShape shape, String name)
  {
    XPropertySet props = Lo.qi(XPropertySet.class, shape);
    try {
      props.setPropertyValue("FillStyle", FillStyle.GRADIENT);
      props.setPropertyValue("FillGradientName", name);
    }
    catch(com.sun.star.lang.IllegalArgumentException e)
    {  System.out.println("\"" + name + "\" is not a recognized gradient color name");  }
    catch(Exception e)
    {  System.out.println("Could not set gradient color to \"" + name + "\"");  }
  }  // end of setGradientColor()



  public static void setGradientColor(XShape shape, 
                           java.awt.Color startColor, java.awt.Color endColor)
  {  setGradientColor(shape, startColor, endColor, 0);  }


  public static void setGradientColor(XShape shape, 
                  java.awt.Color startColor, java.awt.Color endColor, int angle)
  // if angle == 90 then gradient is left --> right side of shape
  {
    Gradient grad = new Gradient(); 
    grad.Style = GradientStyle.LINEAR; 
    grad.StartColor = Lo.getColorInt(startColor); 
    grad.EndColor = Lo.getColorInt(endColor); 
    
    grad.Angle = (short)(angle*10);    // in 1/10 degree units
    grad.Border = 0; 
    grad.XOffset = 0; grad.YOffset = 0; 
    grad.StartIntensity = 100; grad.EndIntensity = 100; 
    grad.StepCount = 10; 

    XPropertySet props = Lo.qi(XPropertySet.class, shape);
    try {
      props.setPropertyValue("FillStyle", FillStyle.GRADIENT);
      props.setPropertyValue("FillGradient", grad);
    }
    catch(Exception e)
    {  System.out.println("Could not set gradient colors");  }
  }  // end of setGradientColor()



  public static void setHatchingColor(XShape shape, String name)
  {
    XPropertySet props = Lo.qi(XPropertySet.class, shape);
    try {
      props.setPropertyValue("FillStyle", FillStyle.HATCH);
      props.setPropertyValue("FillHatchName", name);
    }
    catch(com.sun.star.lang.IllegalArgumentException e)
    {  System.out.println("\"" + name + "\" is not a recognized hatching name");  }
    catch(Exception e)
    {  System.out.println("Could not set hatching color to \"" + name + "\"");  }
  }  // end of setHatchingColor()



  public static void setBitmapColor(XShape shape, String name)
  {
    XPropertySet props = Lo.qi(XPropertySet.class, shape);
    try {
      props.setPropertyValue("FillStyle", FillStyle.BITMAP);
      props.setPropertyValue("FillBitmapName", name);
    }
    catch(com.sun.star.lang.IllegalArgumentException e)
    {  System.out.println("\"" + name + "\" is not a recognized bitmap name");  }
    catch(Exception e)
    {  System.out.println("Could not set bitmap color to \"" + name + "\"");  }
  }  // end of setBitmapColor()



  public static void setBitmapFileColor(XShape shape, String fnm)
  {
    XPropertySet props = Lo.qi(XPropertySet.class, shape);
    try {
      props.setPropertyValue("FillStyle", FillStyle.BITMAP);
      props.setPropertyValue("FillBitmapURL", FileIO.fnmToURL(fnm));
    }
    catch(Exception e)
    {  System.out.println("Could not set bitmap color using \"" + fnm + "\"");  }
  }  // end of setBitmapFileColor()




  public static void setLineStyle(XShape shape, LineStyle style)
  {  Props.setProperty(shape, "LineStyle", style);  }  



  public static void setVisible(XShape shape, boolean isVisible)
  {  Props.setProperty(shape, "Visible", isVisible);  }  



  public static int getRotation(XShape shape)
  { return ((Integer)Props.getProperty(shape, "RotateAngle"))/100;  }


  public static void setRotation(XShape shape, int angle)
  {  Props.setProperty(shape, "RotateAngle", angle*100);  } 

  /* "RotateAngle" is deprecated but is much simpler 
      than the matrix approach, and works correctly 
      for rotations around the center */


  public static HomogenMatrix3 getTransformation(XShape shape)
  /* Returns a transformation matrix, which seems to 
     represent a clockwise rotation:
        cos(t)  sin(t) x
       -sin(t)  cos(t) y
          0       0    1      */
  { return (HomogenMatrix3) Props.getProperty(shape, "Transformation");  } 



  public static void printMatrix(HomogenMatrix3 mat)
  {
    System.out.println("Transformation Matrix:");
    System.out.printf("\t%10.2f\t%10.2f\t%10.2f\n", 
              mat.Line1.Column1, mat.Line1.Column2, mat.Line1.Column3);
    System.out.printf("\t%10.2f\t%10.2f\t%10.2f\n", 
              mat.Line2.Column1, mat.Line2.Column2, mat.Line2.Column3);
    System.out.printf("\t%10.2f\t%10.2f\t%10.2f\n", 
              mat.Line3.Column1, mat.Line3.Column2, mat.Line3.Column3);

    double radAngle = Math.atan2(mat.Line2.Column1, mat.Line1.Column1);
                                      // sin(t), cos(t)
    int currAngle = (int)Math.round( Math.toDegrees(radAngle));
    System.out.println("  Current angle: " + currAngle);
    System.out.println();
    System.out.println();
  }  // end of printMatrix()





 // public static void setLineColor(XShape shape, java.awt.Color c)
 // { Props.setProperty(shape, "LineColor", Lo.getColorInt(c));  } 

  //public static void setLineJoint(XShape shape, LineJoint jointType)
  //{  Props.setProperty(shape, "LineJoint", jointType);  } 

 // public static void setLineThickness(XShape shape, int width)
 // {  Props.setProperty(shape, "LineWidth", width*100);  }  // in 1/100 mm units

 // public static void setFillColor(XShape shape, java.awt.Color c)
 // {  Props.setProperty(shape, "FillColor", Lo.getColorInt(c)); } 


 // public static void setShadow(XShape shape, boolean isShadow)
 // {  Props.setProperty(shape, "Shadow", isShadow);  }  


  //public static void setAnimationEffect(XShape shape, AnimationEffect effectKind)
  //{  Props.setProperty(shape, "Effect", effectKind);  }  

  //public static void setClickAction(XShape shape, ClickAction clickKind)
  // {  Props.setProperty(shape, "OnClick", clickKind); }  

  //public static void setBookmark(XShape shape, String bookmarkName)
  //{  Props.setProperty(shape, "Bookmark", bookmarkName);  } 



/*
  public static void setRotationXXX(XShape shape, int angle)
  // see ObjectTransformationDemo.java for an example
  {
    Point pt = shape.getPosition();    // in 1/100 mm units
    System.out.println("Position: " + pt.X + ", " + pt.Y);


    if (isImage(shape)) {
      System.out.println("Calculate bounded box!");
    }


    Size sz = shape.getSize();    // in 1/100 mm units
    System.out.println("Size: " + sz.Width + ", " + sz.Height);

    int xOffset = pt.X + sz.Width/2;    // of center from origin
    int yOffset = pt.Y + sz.Height/2;

    try {
      XPropertySet props = Lo.qi(XPropertySet.class, shape);

      HomogenMatrix3 mat = (HomogenMatrix3) props.getPropertyValue("Transformation");
      printHMat(mat);

      double radAngle = Math.atan2(mat.Line2.Column1, mat.Line1.Column1);
                                      // sin(t), cos(t)
      int currAngle = (int)Math.round( Math.toDegrees(radAngle));
      System.out.println("Current angle: " + currAngle);

      AffineTransform currMat = new AffineTransform(
                         mat.Line1.Column1, mat.Line2.Column1,
                         mat.Line1.Column2, mat.Line2.Column2,
                         mat.Line1.Column3, mat.Line2.Column3);

      AffineTransform toOrigMat = new AffineTransform();   // move to origin
      toOrigMat.setToTranslation(-xOffset, -yOffset);
      toOrigMat.concatenate(currMat);

      int changeAngle = angle - currAngle;
      AffineTransform rotMat = new AffineTransform();    // rotate by changed angle
      rotMat.setToRotation(Math.PI / 180.0 * changeAngle);
      rotMat.concatenate(toOrigMat);

      double flatMat[] = new double[6];
      rotMat.getMatrix(flatMat);

      mat.Line1.Column1 = flatMat[0];
      mat.Line2.Column1 = flatMat[1];
      mat.Line1.Column2 = flatMat[2];
      mat.Line2.Column2 = flatMat[3];
      mat.Line1.Column3 = flatMat[4];
      mat.Line2.Column3 = flatMat[5];
      props.setPropertyValue("Visible", false);
      props.setPropertyValue("Transformation", mat);


      sz = shape.getSize();    // in 1/100 mm units
      System.out.println("New Size: " + sz.Width + ", " + sz.Height);
      xOffset = pt.X + sz.Width/2;    // of center from origin
      yOffset = pt.Y + sz.Height/2;

      currMat = new AffineTransform(
                         mat.Line1.Column1, mat.Line2.Column1,
                         mat.Line1.Column2, mat.Line2.Column2,
                         mat.Line1.Column3, mat.Line2.Column3);

      AffineTransform fromOrigMat = new AffineTransform();  // move back from origin
      fromOrigMat.setToTranslation(xOffset, yOffset);
      fromOrigMat.concatenate(currMat);

      fromOrigMat.getMatrix(flatMat);

      mat.Line1.Column1 = flatMat[0];
      mat.Line2.Column1 = flatMat[1];
      mat.Line1.Column2 = flatMat[2];
      mat.Line2.Column2 = flatMat[3];
      mat.Line1.Column3 = flatMat[4];
      mat.Line2.Column3 = flatMat[5];
      props.setPropertyValue("Visible", true);
      props.setPropertyValue("Transformation", mat);
    }
    catch (Exception ex) {
      System.out.println("Could not set rotation angle");
    }
  }  // end of setRotationXXX()
*/




  // ==================== draw an image =====================



  public static XShape drawImage(XDrawPage slide, String imFnm)
  {
    Size slideSize = Draw.getSlideSize(slide);  // returned in mm units
    Size imSize = Images.getSize100mm(imFnm);   // in 1/100 mm units
    if (imSize == null) {
      System.out.println("Could not calculate size of " + imFnm);
      return null;
    }
    else {   // center the image on the page
      int imWidth = imSize.Width/100;   // in mm units
      int imHeight = imSize.Height/100;
      int x = (slideSize.Width - imWidth)/2;
      int y = (slideSize.Height - imHeight)/2;
      return drawImage(slide, imFnm, x, y, imWidth, imHeight);
    }
  }  // end of drawImage()



  public static XShape drawImage(XDrawPage slide,
                                   String imFnm, int x, int y)
  // units in mm's
  { Size imSize = Images.getSize100mm(imFnm);   // in 1/100 mm units
    if (imSize == null) {
      System.out.println("Could not calculate size of " + imFnm);
      return null;
    }
    else
      return drawImage(slide, imFnm, x, y, imSize.Width/100, imSize.Height/100);
  }  // end of drawImage()



  public static XShape drawImage(XDrawPage slide,
                    String imFnm, int x, int y, int width, int height)
  // units in mm's
  {
    System.out.println("Adding the picture \"" + imFnm + "\"");
    XShape imShape = addShape(slide, "GraphicObjectShape", 
                                              x, y, width, height);
    setImage(imShape, imFnm);
    setLineStyle(imShape, LineStyle.NONE);
                          // so no border around the image
    return imShape;
  }  // end of drawImage()



  public static void setImage(XShape shape, String imFnm)
  {
    String bitmap = Images.getBitmap(imFnm);
    Props.setProperty(shape, "GraphicURL", bitmap);
                // embed bitmap from image file

    // Props.setProperty(shape, "GraphicURL", FileIO.fnmToURL(imFnm));
                // link to image file
  }  // end of setImage()




  public static XShape drawImageOffset(XDrawPage slide,
                          String imFnm, double xOffset, double yOffset)
  /* insert the specified picture onto the slide page in the doc
     presentation document. Use the supplied (x, y) offsets to locate the
     top-left of the image.
  */
  {
    if ((xOffset < 0) || (xOffset >= 1)) {
      System.out.println("xOffset should be between 0-1; using 0.5");
      xOffset = 0.5;
    }
    if ((yOffset < 0) || (yOffset >= 1)) {
      System.out.println("yOffset should be between 0-1; using 0.5");
      yOffset = 0.5;
    }

    Size slideSize = Draw.getSlideSize(slide);  // returned in mm units
    if (slideSize == null) {
      System.out.println("Image drawing cannot proceed");
      return null;
    }
    int x = (int)Math.round(slideSize.Width * xOffset);   // in mm units
    int y = (int)Math.round(slideSize.Height * yOffset);

    int maxWidth = slideSize.Width - x;
    int maxHeight = slideSize.Height - y;
    Size imSize = Images.calcScale(imFnm, maxWidth, maxHeight);  // in mm units

    return drawImage(slide, imFnm, x, y, imSize.Width, imSize.Height);
  }  // end of drawImageOffset()




  public static boolean isImage(XShape shape)
  {
    return (shape.getShapeType().equals("com.sun.star.drawing.GraphicObjectShape"));
  }  // end of isImage()



  // ----------------------------- form manipulation ---------------------------


  public static XIndexContainer getFormContainer(XDrawPage slide)
  {
    try {
      XFormsSupplier xSuppForms = Lo.qi(XFormsSupplier.class, slide);
      if (xSuppForms == null) {
        System.out.println("Could not access forms supplier");
        return null;
      }

      XNameContainer xFormsCon = xSuppForms.getForms();
      if (xFormsCon == null) {
        System.out.println("Could not access forms container");
        return null;
      }

      XIndexContainer xForms = Lo.qi(XIndexContainer.class, xFormsCon);
      XIndexContainer xForm = Lo.qi(XIndexContainer.class, xForms.getByIndex(0));  // the first form
      return xForm;
    }
    catch(Exception e)
    {  System.out.println("Could not find a form");  }

    return null;
  }  // end of getFormContainer()


  // ----------------------- slide show related --------------------------


  public static XPresentation2 getShow(XComponent doc)
  {
    XPresentationSupplier ps = Lo.qi(XPresentationSupplier.class, doc);
    return Lo.qi(XPresentation2.class, ps.getPresentation());
  }



  public static XSlideShowController getShowController(XPresentation2 show)
  // keep trying to get the slide show controller
  {
    XSlideShowController sc = show.getController();   
       // may return null if executed too quickly after start of show
    int numTries = 1;
    if ((sc == null) && (numTries < 4)) {  // try 3 times
      Lo.delay(1000);  // give the slide show time to start
      numTries++;
      sc = show.getController();
    }
    if (sc == null) 
      System.out.println("Could not obtain slide show controller");
    return sc;
  }  // end of XSlideShowController getShowController()




  public static void waitEnded(XSlideShowController sc)
  /* wait for until the slide is ended, which occurs when
     the user exits the slide show */
  { 
    while (sc.getCurrentSlideIndex() != -1)  // presentation not ended 
      Lo.delay(1000);
    System.out.println("End of presentation detected");
  }



  public static void waitLast(XSlideShowController sc, int delay)
  /* wait for delay milliseconds when the last slide is shown before
     returning */
  {
    int numSlides = sc.getSlideCount();
    // System.out.println("No. of slides: " + numSlides);

    while (sc.getCurrentSlideIndex() < numSlides-1) {  // has not reached last slide?
      // System.out.println("Current slide: " + sc.getCurrentSlideIndex());
      Lo.delay(500);
    }

    // System.out.println("Current slide: " + sc.getCurrentSlideIndex());
    // System.out.println("Next slide: " + sc.getNextSlideIndex());
    Lo.delay(delay);
  }  // end of waitLast()




  public static void setTransition(XDrawPage currSlide, FadeEffect fadeEffect, 
                    AnimationSpeed speed, int change, int duration)
  { try {
      XPropertySet props = Lo.qi(XPropertySet.class, currSlide);
      props.setPropertyValue("Effect", fadeEffect);
      props.setPropertyValue("Speed", speed);
      props.setPropertyValue("Change", change); // see constants at top of this file
      props.setPropertyValue("Duration", duration);  // in seconds
    }
    catch (Exception e) 
    {  System.out.println("Could not set slide transition");  }
  }  // end of setTransition()




  public static XNameContainer buildPlayList(XComponent doc, 
                                     int[] slideIdxs, String customName)
  /* build a named play list container of  slides from doc.
     The name of the play list is customName
  */
  {
    // get a named container for holding the custom play list
    XNameContainer playList = Draw.getPlayList(doc);
    try {
      // create an indexed container for the play list
      XSingleServiceFactory xFactory = Lo.qi(XSingleServiceFactory.class, playList);
      XIndexContainer slidesCon = 
                  Lo.qi(XIndexContainer.class, xFactory.createInstance());

      // container holds slide references whose indicies come from slideIdxs
      System.out.println("Building play list using: ");
      for(int j=0; j < slideIdxs.length; j++) {
        XDrawPage slide = Draw.getSlide(doc, slideIdxs[j]);
        if (slide != null) {
          slidesCon.insertByIndex(j, slide);
          System.out.println("  Slide " + slideIdxs[j]);
        }
      }  

      // store the play list under the custom name
      playList.insertByName(customName, slidesCon);
      System.out.println("Play list stored under the name: " + customName + "\n");
      return playList;
    }
    catch (com.sun.star.uno.Exception e) 
    {  System.out.println("Unable to build play list: " + e);  
       return null;
    }
  }  // end of buildPlayList()




  public static XNameContainer getPlayList(XComponent doc)
  // get a named container for holding custom play lists
  {
    XCustomPresentationSupplier cpSupp = Lo.qi(XCustomPresentationSupplier.class, doc);
    return cpSupp.getCustomPresentations();
  }


  public static XAnimationNode getAnimationNode(XDrawPage slide)
  {
    XAnimationNodeSupplier nodeSupp = Lo.qi(XAnimationNodeSupplier.class, slide);
    return nodeSupp.getAnimationNode();
  }  // end of getAnimationNode()


}  // end of Draw class


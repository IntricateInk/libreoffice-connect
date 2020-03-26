
// Print.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* A growing collection of utility functions to make Office
   easier to use. They are currently divided into the following
   groups:

     * setup printer
     * report printer props
     * print a document
*/

package utils;

import com.sun.star.beans.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;
import com.sun.star.awt.*;

import com.sun.star.document.*;
import com.sun.star.view.*;

// import com.sun.star.uno.Exception;
// import com.sun.star.io.IOException;




public class Print
{


  public static void setDocListener(XPrintableBroadcaster pb)
  // old style listener (?) from 2002
  {
    pb.addPrintableListener( new XPrintableListener()
    {
       public void stateChanged(PrintableStateEvent e) 
       { System.out.println("Print state change: " + printableState(e.State));  }

       public void disposing(com.sun.star.lang.EventObject e)
       {  System.out.println("Disposing of print job: " + e);  }
    });
  }  // end of setDocListener()




  // --------------- setup printer -------------------


  public static void usePrinter(XPrintable xp, String printer)
  {
    if (xp == null) {
      System.out.println("Cannot set printer XPrintable is null");
      return;
    }
    System.out.println("Using printer \"" + printer + "\"");
    xp.setPrinter( Props.makeProps("Name", printer, 
                                    "PaperFormat", PaperFormat.A4) );
    setListener(xp);
  } 


  public static void setListener(XPrintable xp)
  {
    if (xp == null) {
      System.out.println("Cannot set listener; XPrintable is null");
      return;
    }

    XPrintJobBroadcaster pb = Lo.qi(XPrintJobBroadcaster.class, xp);
    if (pb == null) {
      System.out.println("Cannot obtain print job broadcaster");
      return;
    }

    pb.addPrintJobListener( new XPrintJobListener()
    {
       public void printJobEvent(PrintJobEvent e) 
       { System.out.println("Print Job status: " + printableState(e.State));  }

       public void disposing(com.sun.star.lang.EventObject e)
       {  System.out.println("Disposing of print job: " + e);  }
    });
  }  // end of setListener()




  public static String printableState(PrintableState val)
  {
   if (val == PrintableState.JOB_STARTED )
     return "JOB_STARTED";
   else if (val == PrintableState.JOB_COMPLETED )
     return "JOB_COMPLETED";
   else if (val == PrintableState.JOB_SPOOLED )
     return "JOB_SPOOLED";
   else if (val == PrintableState.JOB_ABORTED )
     return "JOB_ABORTED";
   else if (val == PrintableState.JOB_FAILED )
     return "JOB_FAILED";
   else if (val == PrintableState.JOB_SPOOLING_FAILED )
     return "JOB_SPOOLING_FAILED";
   else {
     System.out.println("Unknown printable state");
     return "??";
   }
  }  // end of printableState()



  public static XPropertySet getDocSettings(int docType)
  {
    XPropertySet props = null;
    if (docType == Lo.WRITER)
      props = Lo.createInstanceMSF(XPropertySet.class,
                               "com.sun.star.text.DocumentSettings");
    else if (docType == Lo.IMPRESS)
      props = Lo.createInstanceMSF(XPropertySet.class,
                               "com.sun.star.presentation.DocumentSettings");
    else if (docType == Lo.DRAW)
      props = Lo.createInstanceMSF(XPropertySet.class,
                               "com.sun.star.drawing.DocumentSettings");
    else if (docType == Lo.CALC)
      props = Lo.createInstanceMSF(XPropertySet.class,
                               "com.sun.star.sheet.DocumentSettings");
    else if (docType == Lo.BASE)
      System.out.println("No document settings for a base document");
    else if (docType == Lo.MATH)
      System.out.println("No document settings for a math document");
    else 
      System.out.println("Unknown document type");

    return props;
  }  // end of getDocSettings()



  // ------------------- report printer props -------------------


  public static void reportPrinterProps(XPrintable xp)
  {
    if (xp == null) {
      System.out.println("Cannot report printer props; XPrintable is null");
      return;
    }

    PropertyValue[] printProps = xp.getPrinter();
    if (printProps == null)
      System.out.println("No Printer properties found"); 
    else {
      System.out.println("Printer properties:"); 
      String name;
      for (PropertyValue prop : printProps) {
        name = prop.Name;
        if (name.equals("PaperOrientation"))
          System.out.println("  " + name + ": " + 
                          paperOrientation((PaperOrientation)prop.Value)); 
        else if (name.equals("PaperFormat"))
          System.out.println("  " + name + ": " + 
                         paperFormat((PaperFormat)prop.Value)); 
        else if (name.equals("PaperSize")) {
          Size sz = (Size)prop.Value;
          System.out.println("  " + name + ": (" + sz.Width + ", " + sz.Height + ")"); 
        }
        else
          System.out.println("  " + name + ": " + prop.Value); 
      }
      System.out.println(); 
    }
  }  // end of reportPrinterProps()



  public static String paperOrientation(PaperOrientation val)
  {
   if (val == PaperOrientation.PORTRAIT)
     return "PORTRAIT";
   else if (val == PaperOrientation.LANDSCAPE)
     return "LANDSCAPE";
   else {
     System.out.println("Unknown paper orientation");
     return "??";
   }
  }  // end of paperOrientation()



  public static String paperFormat(PaperFormat val)
  {
   if (val == PaperFormat.A3)
     return "A3";
   else if (val == PaperFormat.A4)
     return "A4";
   else if (val == PaperFormat.A5)
     return "A5";
   else if (val == PaperFormat.B4)
     return "B4";
   else if (val == PaperFormat.B5)
     return "B5";
   else if (val == PaperFormat.LETTER)
     return "LETTER";
   else if (val == PaperFormat.LEGAL)
     return "LEGAL";
   else if (val == PaperFormat.TABLOID)
     return "TABLOID";
   else if (val == PaperFormat.USER)
     return "USER";
   else {
     System.out.println("Unknown paper format");
     return "??";
   }
  }  // end of paperFormat()



  public static String selectionType(SelectionType val)
  {
   if (val == SelectionType.NONE )
     return "NONE";
   else if (val == SelectionType.SINGLE  )
     return "SINGLE";
   else if (val == SelectionType.MULTI  )
     return "MULTI";
   else if (val == SelectionType.RANGE  )
     return "RANGE";
   else {
     System.out.println("Unknown selection type");
     return "??";
   }
  }  // end of selectionType()





  // --------------- print a document -------------------------


  public static boolean isPrintable(int docType)
  { return ((docType == Lo.WRITER) || (docType == Lo.CALC) ||
            (docType == Lo.DRAW) || (docType == Lo.IMPRESS));  }




  public static void print(XPrintable xp)
  {  print(xp, "1-");  }


  public static void print(XPrintable xp, String pagesStr)
  {
    if (xp == null) {
      System.out.println("Cannot print; XPrintable is null");
      return;
    }

    // set the desired pages that will be printed (e.g. "1, 3, 4-7, 9-")
    System.out.println("Print range: " + pagesStr);

    System.out.println("Sending document...");

    PropertyValue[] props = Props.makeProps("Pages", pagesStr,
                                            "Wait", true);  // synchronous
                   // see com.sun.star.view.PrintOptions

    xp.print(props);    // print the document
    System.out.println("Delivered");
  }  // end of print()


}  // end of Print class



// JPrint.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* A growing collection of utility functions to make printing
   easier using the Java print API
    * list printer services
    * find a service
    * quick printing
    * print job listener class: PJWatcher

   For Office printing utilities, see Print.java
*/

package utils;

import java.io.*;
import java.awt.print.*;
import java.util.*;

import javax.print.*;
import javax.print.event.*;
import javax.print.attribute.*;
import javax.print.attribute.standard.*;





public class JPrint
{

  // -------------- list printer services --------------------------


  public static void listServices()
  {  listServices(false);  }


  public static void listServices(boolean showAll)
  {
    PrintService[] psa = PrintServiceLookup.lookupPrintServices(null, null);
    listServices(psa, showAll);
  }


  public static void listServices(PrintService[] psa)
  {  listServices(psa, false);  }


  public static void listServices(PrintService[] psa, boolean showAll)
  {
    if (psa != null && psa.length > 0) {
      System.out.println("\n-------- Print services (" + psa.length + ") ----------\n");
      for (int i = 0; i < psa.length; i++) {
        System.out.println((i+1) + ". \"" + psa[i].getName() + "\""); 
        listService(psa[i], showAll);
      }
      System.out.println("------------------\n");
    }
    else
      System.out.println("No print services found");
  }  // end of listServices()



  public static void listService(PrintService ps)
  {  listService(ps, false);  }



  public static void listService(PrintService ps, boolean showAll)
  {
    Attribute[] attrs = ps.getAttributes().toArray();
    for(Attribute attr : attrs)
      System.out.println("  " + attr.getName() + ":" + attr);

    // print the document types it can print
    System.out.print("  Supported doc types: ");
    DocFlavor[] flavors = ps.getSupportedDocFlavors();
    for (int j = 0; j < flavors.length; j++) {
      // Filter out DocFlavors that have a representation class other
      // than java.io.InputStream.  
      String repclass = flavors[j].getRepresentationClassName();

      if (!repclass.equals("java.io.InputStream"))
        continue;
      System.out.print(" " + flavors[j].getMimeType());
    }
    System.out.println();

    if (showAll) {
      System.out.println("  ----");
      ArrayList<NamedAttribute> attrList = getAttributes(ps);
      Collections.sort(attrList);
      for (NamedAttribute attr : attrList)
        printAttr(ps, attr.getAttribute());
    }
    System.out.println();
  }  // end of listService()


	@SuppressWarnings({ "unchecked", "rawtypes" })
  public static ArrayList<NamedAttribute> getAttributes(PrintService ps) 
  // http://stackoverflow.com/questions/5567709/extended-printer-information-in-java
  // https://docs.oracle.com/javase/7/docs/api/index.html?javax/print/DocFlavor.html
  {
    ArrayList<NamedAttribute> attrList = new ArrayList<NamedAttribute>();

    //get the supported docflavors, categories and attributes
    Class<? extends Attribute>[] cats = 
       (Class<? extends Attribute>[]) ps.getSupportedAttributeCategories();
    DocFlavor[] flavors = ps.getSupportedDocFlavors();
    AttributeSet attributes = ps.getAttributes();

    // get all the available attributes
    for (Class<? extends Attribute> category : cats) {
      for (DocFlavor flavor : flavors) {
        // get the value
        Object value = 
           ps.getSupportedAttributeValues(category, flavor, attributes);
        if (value != null) {
          // if it's a SINGLE attribute...
          if (value instanceof Attribute) {
              Attribute attr = (Attribute) value;
              NamedAttribute na = new NamedAttribute(attr);
              if (!attrList.contains(na))
			          attrList.add(na);
          }
          // if it's a SET of attributes...
          else if (value instanceof Attribute[]) {
            Attribute[] attrs = (Attribute[]) value; // add children
            for(Attribute attr : attrs) {
              NamedAttribute na = new NamedAttribute(attr);
              if (!attrList.contains(na))
			          attrList.add(na);
            }
          }
        }
      }
    }
    return attrList;
  } // end of getAttributes()



  public static void printAttr(PrintService ps, Attribute attr)
  {
    Object value = ps.getDefaultAttributeValue(attr.getCategory());
    System.out.println("  " + attr.getName() + ": " + value);
  }


  public static String[] getPrinterNames()
  { PrintService[] psa = PrintServiceLookup.lookupPrintServices(null, null);
    return getPrinterNames(psa);
  }


  public static String[] getPrinterNames(PrintService[] psa)
  {
    if (psa != null && psa.length > 0) {
      String[] pNames = new String[psa.length];
      for (int i = 0; i < psa.length; i++)
        pNames[i] = psa[i].getName();
      Arrays.sort(pNames);
      return pNames;
    }
    else {
      System.out.println("No print services found");
      return null;
    }
  }  // end of getPrinterNames()



  public static void printNames(String msg, String[] pNames)
  {
    if ((pNames == null) || (pNames.length == 0))
       System.out.println("**No matching printers support " + msg);
    else {
      System.out.println("Printer names that support " + msg + 
                                            " (" + pNames.length + "):");
      for (String pName : pNames)
        System.out.println("  " + pName);
    }
  }  // end of printNames()


  // ---------------- find a service ---------------



  public static String getDefaultPrinterName()
  { PrintService ps = PrintServiceLookup.lookupDefaultPrintService();
    return ps.getName();
  }


  public static void listService(String pName)
  { listService(pName, false);  }


  public static void listService(String pName, boolean showAll)
  {
    PrintService ps = findService(pName);
    System.out.println("\nServices for \"" + pName + "\"");
    listService(ps, showAll);
  }  // end of listService()




  public static String[] findPrinterNames(String str)
  {
    ArrayList<String> nms = new ArrayList<String>();

    String srchStr = str.toLowerCase();

    PrintService[] psa = PrintServiceLookup.lookupPrintServices(null, null);
    String[] pNames = getPrinterNames(psa);

    for (String pName : pNames)
      if (pName.toLowerCase().contains(srchStr))
        nms.add(pName);

    if (nms.size() == 0) {
      System.out.println("Cannot find a printer that matches: \"" + str + "\"");
      return null;
    }
    else {
      Collections.sort(nms);
      return nms.toArray(new String[nms.size()]);
    }
  }  // end of findPrinterNames()



  public static PrintService findService(String pName)
  {
    AttributeSet attrs = new HashAttributeSet();
    attrs.add(new PrinterName(pName, null));
    PrintService[] psa = PrintServiceLookup.lookupPrintServices(null, attrs);

    if ((psa == null) || (psa.length == 0)) {
      System.out.println("No printer service for \"" + pName + "\"");
      return null;
    }
    
    if (psa.length > 1) {
       System.out.println("Multiple matching printer services for \"" + pName + "\"");
       System.out.println("Using first one: \"" + psa[0].getName() + "\"");
    }
    return psa[0];
  }  // end of findService()



  public static PrintService dialogSelect()
  {
    GUI.setLookFeel();
    PrintService psa[] = 
                PrintServiceLookup.lookupPrintServices(null, null);
    PrintService defaultService = 
                PrintServiceLookup.lookupDefaultPrintService();

    PrintRequestAttributeSet attrs = new HashPrintRequestAttributeSet();  // none
    PrintService service = ServiceUI.printDialog(null, 100, 100,
                                      psa, defaultService, null, attrs);
    if (service == null)
      System.out.println("No print service selected");
    return service;
  }  // end of dialogSelect()



  // ------------------- quick printing -----------------------


  public static void dialogPrint(String fnm)
  {
    PrintService ps = dialogSelect();
    if (ps != null)
      printFile(ps, fnm);
  }  // end of dialogPrint()




  public static void printFile(PrintService ps, String fnm)
  {
    if (ps == null) {
      System.out.println("Print service is null");
      return;
    }

    DocPrintJob printJob = ps.createPrintJob();
    try {
      InputStream is = new FileInputStream(fnm);
      DocFlavor flavor = getFlavor(ps, fnm);
      Doc doc = new SimpleDoc(is, flavor, null);
      printJob.print(doc, null);
      is.close();
    }
    catch(Exception e) {
      System.out.println("Unable to print " + fnm);
      System.out.println(e);
    }
  }  // end of printFile()



  public static DocFlavor getFlavor(PrintService ps, String fnm)
  {
     DocFlavor flavor = getFlavorFromFnm(fnm);
     System.out.println("File-based DocFlavor: " + flavor);
     if (!ps.isDocFlavorSupported(flavor)) {
       System.out.println("Not supported by printer; using autosense");
       flavor = DocFlavor.INPUT_STREAM.AUTOSENSE;
     }
     return flavor;
  }  // end of getFlavor()



  public static DocFlavor getFlavorFromFnm(String fnm)
  {
    String ext = fnm.substring(fnm.lastIndexOf('.') + 1).toLowerCase();
    // System.out.println("File extension: " + ext);

    if (ext.equals("gif"))
      return DocFlavor.INPUT_STREAM.GIF;
    else if (ext.equals("jpeg"))
      return DocFlavor.INPUT_STREAM.JPEG;
    else if (ext.equals("jpg"))
      return DocFlavor.INPUT_STREAM.JPEG;
    else if (ext.equals("png")) {
      return DocFlavor.INPUT_STREAM.PNG;
    }
    else if (ext.equals("ps"))
      return DocFlavor.INPUT_STREAM.POSTSCRIPT;
    else if (ext.equals("pdf"))
      return DocFlavor.INPUT_STREAM.PDF;
    else if (ext.equals("txt"))
      return DocFlavor.INPUT_STREAM.TEXT_PLAIN_HOST;
    else {      // try to determine flavor from file content
      // System.out.println("Using autosense");
      return DocFlavor.INPUT_STREAM.AUTOSENSE;
    }
  }  // end of getFlavorFromFnm



  public static void print(String fnm)
  { PrintService ps = PrintServiceLookup.lookupDefaultPrintService();
    printFile(ps, fnm);
  }  

  public static void print(String pName, String fnm)
  {  printFile(findService(pName), fnm);  }  



  public static void printMonitorFile(PrintService ps, String fnm)
  {
    if (ps == null) {
      System.out.println("Print service is null");
      return;
    }

    DocPrintJob printJob = ps.createPrintJob();
    printJob.addPrintJobListener(new PJWatcher());

    try {
      InputStream is = new FileInputStream(fnm);
      DocFlavor flavor = getFlavor(ps, fnm);
      Doc doc = new SimpleDoc(is, flavor, null);
      printJob.print(doc, null);
      is.close();
    }
    catch(Exception e) {
      System.out.println("Unable to print " + fnm);
      System.out.println(e);
    }
  }  // end of printMonitorFile()



  // -----------------------------------

  private static class PJWatcher implements PrintJobListener
  {
    public void printDataTransferCompleted(PrintJobEvent pje)
    {  System.out.println(" >> Data transferred to printer"); }

    public void printJobCanceled(PrintJobEvent pje)
    {  System.out.println(" >> Print job was cancelled");  }

    public void printJobCompleted(PrintJobEvent pje)
    {  System.out.println(" >> Print job completed successfully");  }

    public void printJobFailed(PrintJobEvent pje)
    {  System.out.println(" >> Print job failed");  }

    public void printJobNoMoreEvents(PrintJobEvent pje)
    {  System.out.println(" >> No more events will be delivered");  }

    public void printJobRequiresAttention(PrintJobEvent pje)
    {  System.out.println(" >> Print job needs attention"); }

  }  // end of PJWatcher class


}  // end of JPrint class


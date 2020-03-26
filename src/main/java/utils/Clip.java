
// Clip.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Office Clipboard support functions:
       * use the system clipboard

       * set/get text
           -- uses my TextTransferable class

       * get/set a BufferedImage
           -- uses my ImageTransferable class

       * store a file on the clipboard (as bytes)
           -- uses my FileTransferable class

       * flavor functions

       * get functions that try to retrieve data of a
         particular type from the clipboard

   See JClip.java for clipboard functions using only Java,
   not the Office API.

   A good clipboard tool: ClCl
      http://www.nakka.com/soft/clcl/index_eng.html
*/

package utils;

import java.io.*;
import java.awt.image.*;
import javax.imageio.*;


import com.sun.star.uno.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;

import com.sun.star.datatransfer.*;
import com.sun.star.datatransfer.clipboard.*;
import com.sun.star.lang.EventObject;




public class Clip
{
  private static final int MAX_TRIES = 3;


  private static XSystemClipboard cb = null;
        // used to store clipboard ref, so only one is created
        // by the methods here



  public static XSystemClipboard getClip()
  { 
    if (cb == null) {
      cb = SystemClipboard.create(Lo.getContext());  
/*
      cb.addClipboardListener( new XClipboardListener()
      {
        public void disposing(EventObject e) { }
     
        public void changedContents(ClipboardEvent e)
        { System.out.println(">>Clipboard has been updated");  }
      });
*/
    }
    return cb;
  }  // end of getClip()



  public static boolean addContents(XTransferable trf)
  {
    int i = 0;
    while (i < MAX_TRIES) {
      try {
        getClip().setContents(trf, null);
        return true;
      }
      catch (IllegalStateException e) {
        System.out.println("Problem accessing clipboard...");
        Lo.wait(50);
      }
      i++;
    }
    System.out.println("Unable to add contents");
    return false;
  }  // end of addContents()




  public static boolean setText(String str)
  {  return addContents( new TextTransferable(str));  }


  public static String getText()
  {  return (String) getData("text/plain;charset=utf-16");  }



  public static Object getData(String mimeStr)
  /* return the data associated with the specified mime type
     string, or return null */
  {
    XTransferable trf = getClip().getContents();
    if (trf == null) {
      System.out.println("No transferable found");
      return null;
    }

    try {
      DataFlavor df = findFlavor(trf, mimeStr);
      if (df != null)
        return trf.getTransferData(df);
      else
        System.out.println("Mime type \"" + mimeStr + "\" not found");
    }
    catch (com.sun.star.uno.Exception e) {
      System.out.println("Could not read clipboard: " + e);
    }
    return null;
  }  // end of getData()




  public static boolean setImage(BufferedImage im)
  {  return addContents(new ImageTransferable(im));  }


  public static BufferedImage getImage()
  {
    XTransferable trf = getClip().getContents();
    if (trf == null) {
      System.out.println("No transferable found");
      return null;
    }

    DataFlavor df = findImageFlavor(trf);
    if (df == null)
      return null;

    try {
      return Images.bytes2im( (byte[])trf.getTransferData(df) );
    }
    catch (com.sun.star.uno.Exception e) {
      System.out.println("Could not retrieve image from clipboard: " + e);
      return null;
    }
  }  // end of getImage()


  public static BufferedImage readImage()
  // included for backward compatibility with old code
  {  return getImage();  }


  public static boolean setFile(String fnm)
  {  return addContents(new FileTransferable(fnm));  }


  public static byte[] getFile(String fnm)
  {   
    String mimeStr = Info.getMIMEType(fnm);
    System.out.println("MIME Type: " + mimeStr);
    return (byte[]) getData(mimeStr);
  }





  // ----------------- flavor-related functions ---------------------


  public static DataFlavor findFlavor(XTransferable trf, String mimeStr)
  /* search through the flavors associated with the transferable
     looking for the specified mime type string
  */
  {
    DataFlavor[] dfs = trf.getTransferDataFlavors();
    for (int i = 0; i < dfs.length; i++) {
      if (dfs[i].MimeType.startsWith(mimeStr)) {
        // System.out.println("Mime type is: \"" + dfs[i].MimeType + "\"");
        return dfs[i];
      }
    }
    System.out.println("Clip does not support mime type: " + mimeStr);
    return null;
  }  // end of findFlavor()




  public static DataFlavor findImageFlavor(XTransferable trf)
  // find the first image-related flavor for this transferable, or null
  {
    DataFlavor[] dfs = trf.getTransferDataFlavors();
    for (int i = 0; i < dfs.length; i++) {
      if (Info.isImageMime(dfs[i].MimeType)) {
        System.out.println("Found mime type: \"" + dfs[i].MimeType + "\"");
        return dfs[i];
      }
    }
    System.out.println("Clip does not support an image mime type");
    return null;
  }  // end of findImageFlavor()



  public static String getFirstMime(XTransferable trf)
  /* return the MIME type string in the first flavor for
     this transferable 
  */
  {
    DataFlavor[] dfs = trf.getTransferDataFlavors();
    if (dfs.length > 0) {
      System.out.println("Using first Mime type: " + dfs[0].MimeType);
      return dfs[0].MimeType;
    }
    else {
      System.out.println("No Mime type found");
      return null;
    }
  }  // end of getFirstMime()




  public static void listFlavors()
  /* print the MIME type strings for all the flavors supported
     by the current clip */
  {
    XTransferable trf = getClip().getContents();
    if (trf == null)
      System.out.println("No transferable found");
    else
      listFlavors(trf);
  }  // end of listFlavors()


  public static void listFlavors(XTransferable trf)
  /* print the MIME type strings for all the flavors supported
     by the transferable */
  {
    DataFlavor[] dfs = trf.getTransferDataFlavors();
    System.out.println("No of flavors: " + dfs.length);
    for (int i = 0; i < dfs.length; i++)
      System.out.println((i+1) + ". " + dfs[i].MimeType);
    System.out.println();
  }  // end of listFlavors()



  // ----------- get data from the clipboard  -------------


  public static String getHTML()
  {  
    byte[] data = (byte[]) getData("text/html");
    if (data == null)
      return null;
    else
      return new String(data);
  }  // end of getHTML()



  public static String getRTF()
  // text-based text format used by MS
  {  
    byte[] data = (byte[]) getData("text/richtext");
    if (data == null)
      return null;
    else
      return new String(data);
  }  // end of getRTF()



  public static String getSylk()
  // text-based spreadsheet format used by MS
  {  
    byte[] data = (byte[]) getData(
        "application/x-openoffice-sylk;windows_formatname=\"Sylk\"");
    if (data == null)
      return null;
    else
      return new String(data);
  }  // end of getSylk()




  public static String getXMLDraw()
  // a 'flat' XML version of the draw/slide copy
  {  
    byte[] data = (byte[]) getData(
        "application/x-openoffice-drawing;windows_formatname=\"Drawing Format\"");
    if (data == null)
      return null;
    else
      return new String(data);
  }  // end of getXMLDraw()



  public static byte[] getEmbedSource()
  // fragment in binary ODF format 
  {
    return (byte[]) getData(
        "application/x-openoffice-embed-source-xml;windows_formatname=\"Star Embed Source (XML)\"");
  }  // end of getEmbedSource()



  public static byte[] getBitmap()
  // bitmap image, suitable for use by BitmapTable service;
  // see Images.getBitmap()
  {
    return (byte[]) getData(
        "application/x-openoffice-bitmap;windows_formatname=\"Bitmap\"");
  }  // end of getBitmap()



}  // end of Clip class


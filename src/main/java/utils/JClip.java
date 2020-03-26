
// JClip.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Support functions for the Java's clipboard API:
       * use the system clipboard

       * set/get text
            -- uses Java's DataFlavor.stringFlavor

       * get/set a BufferedImage
            -- uses Java's DataFlavor.imageFlavor

       * get/set a 2D array
            -- uses my JArrayTransferable class
            -- useful for Calc and Base clips

       * get/set an ArrayList
            -- uses my JListTransferable class


   See Clip.java for clipboard functions using the Office API.

   A good clipboard tool: ClCl
      http://www.nakka.com/soft/clcl/index_eng.html
*/

package utils;

import java.io.*;
import java.util.*;
import java.awt.*;
import java.awt.image.*;
import java.awt.datatransfer.*;



public class JClip
{
  private static final int MAX_TRIES = 3;


  // flavors for my 2D array and ArrayList data types
  private static final DataFlavor ARRAY_DF = 
                             new DataFlavor(Object[][].class, "2D Object Array");
  private static final DataFlavor LIST_DF = 
                             new DataFlavor(ArrayList.class, "List");


  private static Clipboard cb = null;
        // used to store clipboard ref, so only one is created
        // by the methods here



  public static Clipboard getClip()
  { 
   if (cb == null) {
      cb = Toolkit.getDefaultToolkit().getSystemClipboard();
/*
      cb.addFlavorListener( new FlavorListener() {
        public void flavorsChanged(FlavorEvent e)
        {  System.out.println(">>Flavor change detected");  }
      });
*/
    }
    return cb;
  }


  public static boolean addContents(Transferable trf)
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
  {  return addContents( new StringSelection(str));  }



  public static String getText()
  {
    Transferable trf = getClip().getContents(null);
    try {
      if (trf != null && trf.isDataFlavorSupported(DataFlavor.stringFlavor))
        return (String) trf.getTransferData(DataFlavor.stringFlavor);
    } 
    catch (UnsupportedFlavorException e) 
    {  System.out.println(e); } 
    catch (IOException e)
    { System.out.println(e); }
    return null;
  }  // end of getText()




  public static boolean setImage(BufferedImage im)
  {  return addContents(new JImageTransferable(im));  }



  public static BufferedImage getImage()
  {
    Transferable trf = getClip().getContents(null);
    if (trf != null && trf.isDataFlavorSupported(DataFlavor.imageFlavor)) {
      try {
        return (BufferedImage) trf.getTransferData(DataFlavor.imageFlavor);
      } 
      catch (Exception e) 
      {  System.out.println(e); }
    }
    return null;
  }  // end of getImage()




  public static boolean setArray(Object[][] vals)
  {  return addContents(new JArrayTransferable(vals));  }



  public static Object[][] getArray()
  {
    Transferable trf = getClip().getContents(null);
    if (trf != null && trf.isDataFlavorSupported(ARRAY_DF)) {
      try {
        return (Object[][]) trf.getTransferData(ARRAY_DF);
      } 
      catch (Exception e) 
      {  System.out.println(e); }
    }
    return null;
  }  // end of getArray()



  public static boolean setList(ArrayList vals)
  {  return addContents(new JListTransferable(vals));  }



  public static ArrayList getList()
  {
    Transferable trf = getClip().getContents(null);
    if (trf != null && trf.isDataFlavorSupported(LIST_DF)) {
      try {   // unchecked cast
        return (ArrayList) trf.getTransferData(LIST_DF);
      } 
      catch (Exception e) 
      {  System.out.println(e); }
    }
    return null;
  }  // end of getList()





  public static void listFlavors()
  /* print the MIME type strings for all the flavors supported
     by the current clip */
  {
    Transferable trf = getClip().getContents(null);
    if (trf == null)
      System.out.println("No transferable found");
    else
      listFlavors(trf);
  }  // end of listFlavors()


  public static void listFlavors(Transferable trf)
  /* print the MIME type strings for all the flavors supported
     by the transferable */
  {
    DataFlavor[] dfs = trf.getTransferDataFlavors();
    System.out.println("No of flavors: " + dfs.length);
    for (int i = 0; i < dfs.length; i++)
      System.out.println((i+1) + ". " + dfs[i].getMimeType());
    System.out.println();
  }  // end of listFlavors()

}  // end of JClip class

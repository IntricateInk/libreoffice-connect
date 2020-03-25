
// Gallery.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/*  Methods for accessing and reporting on the Gallery.

*/

package utils;

import com.sun.star.gallery.*;
import com.sun.star.graphic.*;


public class Gallery
{


  public static void reportGallerys()
  {
    XGalleryThemeProvider gtp = Lo.createInstanceMCF(XGalleryThemeProvider.class, 
                                          "com.sun.star.gallery.GalleryThemeProvider");
    String[] themes = gtp.getElementNames();
    System.out.println("No. of themes: " + themes.length);
    for (String theme : themes) {
      // System.out.println("  " + theme);
      try {
        XGalleryTheme gallery = Lo.qi(XGalleryTheme.class, gtp.getByName(theme));
        System.out.println("  \"" + gallery.getName() + "\" (" + gallery.getCount() + ")");
      }
      catch(com.sun.star.uno.Exception e)
      {  System.out.println("Could not access gallery for " + theme);  }
    }
  }  // end of reportGallerys()




  public static XGalleryTheme getGallery(String name)
  {
    XGalleryThemeProvider gtp = Lo.createInstanceMCF(XGalleryThemeProvider.class, 
                                          "com.sun.star.gallery.GalleryThemeProvider");
    try {
      return Lo.qi(XGalleryTheme.class, gtp.getByName(name));
    }
    catch(com.sun.star.uno.Exception e)
    {  System.out.println("Could not access gallery for " + name);  
       return null;
    }
  }  // end of getGallery()



  public static void reportGalleryItems(String galleryName)
  {
    XGalleryTheme gallery = getGallery(galleryName);
    if (gallery == null)
      System.out.println("Could not find the gallery: " + galleryName);  
    else {
      int numPics = gallery.getCount();
      System.out.println("Gallery: \"" + gallery.getName() + "\" (" + numPics + ")");
      for(int i=0; i < numPics; i++) {
        try {
          XGalleryItem item = Lo.qi(XGalleryItem.class, gallery.getByIndex(i));
          String URL = (String)Props.getProperty(item, "URL");
          System.out.println("  " + FileIO.getFnm( FileIO.URI2Path(URL)));
        }
        catch(com.sun.star.uno.Exception e)
        {  System.out.println("Could not access gallery item " + i);  }
      }
    }
  }  // end of reportGalleryItems()



  public static XGalleryItem findGalleryItem(String galleryName, String itemNm)
  {
    String nm = itemNm.toLowerCase();

    XGalleryTheme gallery = getGallery(galleryName);
    if (gallery == null){
      System.out.println("Could not find the gallery: " + galleryName); 
      return null;
    }
 
    int numPics = gallery.getCount();
    System.out.println("Searching gallery " + gallery.getName() + " for \"" + itemNm + "\"");
    for(int i=0; i < numPics; i++) {
      try {
        XGalleryItem item = Lo.qi(XGalleryItem.class, gallery.getByIndex(i));
        String URL = (String)Props.getProperty(item, "URL");
        String fnm = FileIO.getFnm( FileIO.URI2Path(URL));
        if (fnm.toLowerCase().contains(nm)) {
          System.out.println("Found matching item: " + fnm);
          return item;
        }
      }
      catch(com.sun.star.uno.Exception e)
      {  System.out.println("Could not access gallery item " + i);  }
    }
    System.out.println("No match found");
    return null;
  }  // end of findGalleryItem()



  public static void reportGalleryItem(XGalleryItem item)
  {
    if (item == null)
      System.out.println("Gallery item is null");  
    else {
      System.out.println("Gallery item information:");
      String URL = (String)Props.getProperty(item, "URL");
      String path = FileIO.URI2Path(URL);
      System.out.println("  Fnm: \"" + FileIO.getFnm(path) + "\"");
      System.out.println("  Path: \"" + path + "\"");
      System.out.println("  Title: \"" + Props.getProperty(item, "Title") + "\"");
      byte itemType = (Byte)Props.getProperty(item, "GalleryItemType");
      System.out.println("  Type: " + getItemTypeStr(itemType));
    }
  }  // end of reportGalleryItem()



  public static String getItemTypeStr(byte itemType)
  {
    switch (itemType)  {
      case GalleryItemType.EMPTY:
        return "empty";
      case GalleryItemType.GRAPHIC:
        return "graphic";
      case GalleryItemType.MEDIA:
        return "media";
      case GalleryItemType.DRAWING:
        return "drawing";
      default: 
        return "??";
    } 
  }  // end of getItemTypeStr()



  public static String getGalleryPath(XGalleryItem item)
  {
    if (item == null) {
      System.out.println("Gallery item is null");  
      return null;
    }
    else {
      String URL = (String)Props.getProperty(item, "URL");
      return FileIO.URI2Path(URL);
    }
  } // end of getGalleryPath()



  public static XGraphic getGalleryGraphic(XGalleryItem item)
  {
    if (item == null) {
      System.out.println("Gallery item is null");  
      return null;
    }
    else
      return (XGraphic)Props.getProperty(item, "Graphic");
  } // end of getGalleryGraphic()

}  // end of Gallery class

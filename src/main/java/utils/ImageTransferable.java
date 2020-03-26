
// ImageTransferable.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* An Office transferable for reading/writing a BufferedImage 
   to the clipboard.

   The Mime type is Office BMP, and the data is stored
   as a byte array.

   Based on code example in the OO Developer's Guide.

   See JImageTransferable for a non-Office, Java version.
*/

package utils;

import java.io.*;
import java.awt.image.*;
import javax.imageio.*;

import com.sun.star.datatransfer.*;
import com.sun.star.uno.*;



public class ImageTransferable implements XTransferable
{
  private static final String BITMAP_CLIP = 
        "application/x-openoffice-bitmap;windows_formatname=\"Bitmap\"";

  private byte[] imBytes;   // see Dev Guide, p.685

  public ImageTransferable(BufferedImage im)
  {  imBytes = Images.im2bytes(im);   }



  public Object getTransferData(DataFlavor df) 
                               throws UnsupportedFlavorException
  {
    if (!df.MimeType.equalsIgnoreCase(BITMAP_CLIP))
      throw new UnsupportedFlavorException();

    return imBytes;
  }  // end of getTransferData()




  public DataFlavor[] getTransferDataFlavors()
  {
    DataFlavor[] dfs = new DataFlavor[1];
    dfs[0] = new DataFlavor(BITMAP_CLIP, "Bitmap", 
                               new Type(byte[].class));  // see Dev Guide p.685
    return dfs;
  }


  public boolean isDataFlavorSupported(DataFlavor df)
  {  return df.MimeType.equalsIgnoreCase(BITMAP_CLIP);  }

}  // end of ImageTransferable class


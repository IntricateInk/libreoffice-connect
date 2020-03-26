
// JImageTransferable.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Use Java's clipboard API to store a Java Image instance on
   the clipboard. The mime type is DataFlavor.imageFlavor

   See ImageTransferable for a Office version.
*/

package utils;

import java.io.*;
import java.awt.*;
import java.awt.image.*;
import java.awt.datatransfer.*;



public class JImageTransferable implements Transferable 
{
  private Image im;

  public JImageTransferable(Image im) 
  {  this.im = im;  }


  public Object getTransferData(DataFlavor df)
                   throws UnsupportedFlavorException, IOException 
  { if (df.equals(DataFlavor.imageFlavor) && im != null)
      return im;
    else
      throw new UnsupportedFlavorException(df);
  }


  public DataFlavor[] getTransferDataFlavors() 
  { DataFlavor[] dfs = new DataFlavor[1];
    dfs[0] = DataFlavor.imageFlavor;
    return dfs;
  }


  public boolean isDataFlavorSupported(DataFlavor df) 
  {
    DataFlavor[] dfs = getTransferDataFlavors();
    for (int i = 0; i < dfs.length; i++) {
      if (df.equals(dfs[i]))
        return true;
    }
    return false;
  }

}  // end of JImageTransferable class

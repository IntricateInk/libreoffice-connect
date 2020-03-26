
// JArrayTransferable.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Use Java's clipboard API to store a 2D array of Objects
   on the clipboard.

   Useful for clipped data from spreadsheets and database tables.
*/

package utils;

import java.io.*;
import java.awt.*;
import java.awt.image.*;
import java.awt.datatransfer.*;



public class JArrayTransferable implements Transferable 
{

  private Object[][] vals;
  private DataFlavor arrDF;

  public JArrayTransferable(Object[][] vals) 
  {  this.vals = vals;  
     arrDF = new DataFlavor(Object[][].class, "2D Object Array");
  }


  public Object getTransferData(DataFlavor df)
                   throws UnsupportedFlavorException, IOException 
  { if (df.equals(arrDF) && vals != null)
      return vals;
    else
      throw new UnsupportedFlavorException(df);
  }


  public DataFlavor[] getTransferDataFlavors() 
  { DataFlavor[] dfs = new DataFlavor[1];
    dfs[0] = arrDF;
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

}  // end of JArrayTransferable class

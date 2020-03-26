
// TextTransferable.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

// An Office transferable for unicode text

package utils;

import com.sun.star.datatransfer.*;
import com.sun.star.uno.Type;


public class TextTransferable implements XTransferable
{
  private final String UNICODE_MIMETYPE = "text/plain;charset=utf-16";

  private String text;


  public TextTransferable(String s)
  {  text = s;  }


  public Object getTransferData(DataFlavor df) 
                                  throws UnsupportedFlavorException
  { if (!df.MimeType.equalsIgnoreCase(UNICODE_MIMETYPE))
      throw new UnsupportedFlavorException();
    return text;
  }


  public DataFlavor[] getTransferDataFlavors()
  { DataFlavor[] dfs = new DataFlavor[1];
    dfs[0] = new DataFlavor(UNICODE_MIMETYPE, "Unicode Text",
                                                  new Type(String.class));
    return dfs;
  }


  public boolean isDataFlavorSupported(DataFlavor df)
  {  return df.MimeType.equalsIgnoreCase(UNICODE_MIMETYPE);  }

}  // end of TextTransferable class

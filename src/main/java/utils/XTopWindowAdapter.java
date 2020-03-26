
// XTopWindowAdapter.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

package utils;

import com.sun.star.lang.*;
import com.sun.star.awt.*;



public class XTopWindowAdapter implements XTopWindowListener
{

  public XTopWindowAdapter(){}

  public void windowOpened(EventObject event){}

  public void windowActivated(EventObject event){}

  public void windowDeactivated(EventObject event){}

  public void windowMinimized(EventObject event) {}

  public void windowNormalized(EventObject event){}

  public void windowClosing(EventObject event){}

  public void windowClosed(EventObject event){}

  public void disposing(EventObject event){}

}  // end of XTopWindowAdapter class
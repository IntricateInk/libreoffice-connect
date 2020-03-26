
// ItemInterceptor.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, August 2016


/* Intercept dispatches sent to the toolbar item, and
   call ToolbarItemListener.clicked() with the dispatch 
   details.

   Based on:
     https://forum.openoffice.org/en/forum/viewtopic.php?f=44&t=47662
     https://forum.openoffice.org/en/forum/viewtopic.php?f=25&t=44291
*/

package utils;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;
import com.sun.star.util.*;



public class ItemInterceptor implements XDispatchProviderInterceptor, XDispatch
{
  private XDispatchProvider slaveDP, masterDP;
             // pointers to next and previous dispatch providers in chain

  private ToolbarItemListener viewer;   // object sent dispatch info
  private String itemName;  // toolbar item name
  private String cmd;       // toolbar item's dispatch command




  public ItemInterceptor(ToolbarItemListener v, String itemName) 
  {  viewer = v;  
     this.itemName = itemName;
     cmd = Lo.makeUnoCmd(itemName);
  }




  public XDispatch[] queryDispatches(DispatchDescriptor[] descrs) 
  { 
    int count = descrs.length; 
    XDispatch[] xDispatch = new XDispatch[count]; 

    for (int i = 0; i < count; i++) 
      xDispatch[i] = queryDispatch(descrs[i].FeatureURL, descrs[i].FrameName, 
                                   descrs[i].SearchFlags); 
    return xDispatch; 
  }  // end of queryDispatches()


/*
  public XDispatch[] queryDispatches(DispatchDescriptor[] descrs)
  {
    if (slaveDP != null)
      return slaveDP.queryDispatches(descrs); 
    else
      return null;
  }
*/


  public XDispatch queryDispatch(URL cmdURL, String target, int srchFlags)
  /* intercept command URLs --
       if the command is the toolbar item command then use this object,
       otherwise ignore the command 
  */
  {
    // System.out.println("queryDispatch: " + cmdURL.Complete);

    if (cmdURL.Complete.equalsIgnoreCase(cmd))  {
      System.out.println(itemName + " seen"); 
      return this;   // this will cause dispatch() to be called
    }

    if (slaveDP != null)
      return slaveDP.queryDispatch(cmdURL, target, srchFlags); 
        // pass command to next interceptor in list
    else
      return null;
  }  // end of queryDispatch()



  public void setMasterDispatchProvider(XDispatchProvider dp)
  {  masterDP = dp;  }

  public void setSlaveDispatchProvider(XDispatchProvider dp)
  {  slaveDP = dp;  }

  public XDispatchProvider getMasterDispatchProvider()
  {  return masterDP;  }

  public XDispatchProvider getSlaveDispatchProvider()
  {  return slaveDP;  }


  // ----------------------- XDispatch methods ---------------------
  

   public void dispatch(URL cmdURL, PropertyValue[] props) 
   // pass command details to the viewer
   {  viewer.clicked(itemName, cmdURL, props);  }


   public void addStatusListener(XStatusListener status, URL url) { }

   public void removeStatusListener(XStatusListener status, URL url) { }


}  // end of ItemInterceptor class

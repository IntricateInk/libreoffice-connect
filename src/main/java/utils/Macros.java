package utils;

// Macros.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Nov. 2016

/* A growing collection of utility functions to make Office
   easier to use. They are currently divided into the following
   groups:

     * execute macros
     * execute LibreLogo
     * list macro script names
     * use XRayTool
     * macro security
     * macros and events
*/


import java.io.*;
import java.util.*;

import com.sun.star.lang.*;
import com.sun.star.container.*;
import com.sun.star.uno.*;
import com.sun.star.beans.*;
import com.sun.star.frame.*;
import com.sun.star.document.*;

import com.sun.star.script.*;
import com.sun.star.script.browse.*;
import com.sun.star.script.provider.*;

import com.sun.star.uno.Exception;
import com.sun.star.io.IOException;
import utils.Info;
import utils.Lo;
import utils.Props;


public class Macros
{

  // simple macro security level words
  // matches those values used in Tools > Options> Security; Macro Security Button dialog
  public static final int LOW = 0;
  public static final int MEDIUM = 1;
  public static final int HIGH = 2;
  public static final int VERY_HIGH = 3;


  // ----------- execute macros ---------------------

  public static Object execute(String macroName, String language, String location)
  {  return execute(macroName, null, language, location);   }



  public static Object execute(String macroName, Object[] params, 
                                         String language, String location)
  {
    if (!isMacroLanguage(language)) {
      System.out.println("\"" + language + "\" is not a macro language name");
      return null;
    }

    try {
      /* deprecated approach
         XScriptProviderFactory spFactory =  Lo.createInstanceMCF( 
                    XScriptProviderFactory.class,
                    "com.sun.star.script.provider.MasterScriptProviderFactory");
      */
      XComponentContext xcc = Lo.getContext();
      XScriptProviderFactory spFactory =  Lo.qi(XScriptProviderFactory.class, 
             xcc.getValueByName(
                "/singletons/com.sun.star.script.provider.theMasterScriptProviderFactory"));

      XScriptProvider sp = spFactory.createScriptProvider(""); 
      XScript xScript = sp.getScript("vnd.sun.star.script:" + macroName + 
                               "?language=" + language + "&location=" + location); 

      // minimal inout/out parameters
      short[][] outParamIndex = { { 0 } };
      Object[][] outParam = { { null } };
      return xScript.invoke(params, outParamIndex, outParam); 
    } 
    catch (Exception e) { 
      System.out.println("Could not execute macro " + macroName + ": " + e); 
      return null;
    } 
  }  // end of execute()



  public static boolean isMacroLanguage(String language)
  {
    return ( language.equals("Basic") || language.equals("BeanShell") ||
             language.equals("JavaScript") || language.equals("Java") ||
             language.equals("Python") );
  } // end of isMacroLanguage



  // ------------ execute LibreLogo ------------------------


  public static Object executeLogoCmds(String cmdsStr)
  {  
    Object[] params = new String[2];
    params[0] = "";    // based on looking at commandline() in LibreLogo.py
    params[1] = cmdsStr;
    return execute("LibreLogo/LibreLogo.py$commandline", 
                                              params,  "Python", "share");  
  }


  public static Object executeLogo(String cmd)
  {  return execute("LibreLogo/LibreLogo.py$"+cmd, "Python", "share");   }
   // <Office>\share\Scripts\python\LibreLogo
   /*  cmd can be "run", "stop", 
       "home", goforward, gobackward, left, right,
       clearscreen, commandline (which is used by executeLogoCmds())
   */



  // ----------------- list macro script names ------------------------


  public static ArrayList<String> getScripts()
  {
    ArrayList<String> scripts = new ArrayList<>();

    XComponentContext xcc = Lo.getContext();
    XBrowseNodeFactory bnf = Lo.qi(XBrowseNodeFactory.class, 
             xcc.getValueByName(
                "/singletons/com.sun.star.script.browse.theBrowseNodeFactory"));
    XBrowseNode rootNode = Lo.qi(XBrowseNode.class,
           bnf.createView( BrowseNodeFactoryViewTypes.MACROORGANIZER) );  // for scripts

    XBrowseNode[] typeNodes = rootNode.getChildNodes();
    for(int i=0; i < typeNodes.length; i++) {
      XBrowseNode typeNode = typeNodes[i];  
      XBrowseNode[] libraryNodes = typeNode.getChildNodes();
      for(int j=0; j < libraryNodes.length; j++)
        getLibScripts(libraryNodes[j], 0, typeNode.getName(), scripts);
    }
    System.out.println();
    return scripts;
  } // end of getScripts()


  public static void getLibScripts(XBrowseNode browseNode, int level,
                                  String path, ArrayList<String> scripts) 
  {
    XBrowseNode[] scriptNodes = browseNode.getChildNodes();
    if ((scriptNodes.length == 0) && (level > 1))  // not a top-level library
      System.out.println("No scripts in " + path);
    for(int i=0; i < scriptNodes.length; i++) {
      XBrowseNode scriptNode = scriptNodes[i];  
      if (scriptNode.getType() == BrowseNodeTypes.SCRIPT) {
        XPropertySet props = Lo.qi(XPropertySet.class, scriptNode);
        if (props != null) {
          try {
            scripts.add((String)props.getPropertyValue("URI"));
            // XScript xScript = scriptProvider.getScript(uri);
          }
          catch(com.sun.star.uno.Exception e) 
          { System.out.println(e); }
        }
        else 
          System.out.println("No props for " + scriptNode.getName());
      }
      else if (scriptNode.getType() == BrowseNodeTypes.CONTAINER)
        getLibScripts(scriptNode, level+1, path + ">" + scriptNode.getName(), scripts);
      else
        System.out.println("Unknown node type");
    }
  }  // end of getLibScripts()



  public static ArrayList<String> getLangScripts(String lang)
  {
    if (!isMacroLanguage(lang)) {
      System.out.println("Not a Macro language; try \"Java\"");
      return null;
    }
    ArrayList<String> fScripts = new ArrayList<>();
    ArrayList<String> scriptURIs = getScripts();
    for(String scriptURI : scriptURIs)
      if (scriptURI.contains("language=" + lang + "&"))  // to avoid JavaScript
        fScripts.add(scriptURI);
    return fScripts;
  } // end of getLangScripts()



  public static ArrayList<String> findScripts(String substr)
  {
    String sub = substr.toLowerCase();
    ArrayList<String> fScripts = new ArrayList<>();
    ArrayList<String> scriptURIs = getScripts();
    for(String scriptURI : scriptURIs)
      if (scriptURI.toLowerCase().contains(sub))
        fScripts.add(scriptURI);
    return fScripts;
  } // end of findScripts()



  // ------------------- use XRayTool -----------------------------
  // available from ??

  public static XScript loadXRay()
  {
    try {
      XScriptProviderFactory spFactory = Lo.createInstanceMCF(
                              XScriptProviderFactory.class,
                              "com.sun.star.script.provider.MasterScriptProviderFactory");

      XScriptProvider provider = spFactory.createScriptProvider("");
      return provider.getScript(
        "vnd.sun.star.script:XrayTool._Main.Xray?language=Basic&location=application");
    } 
    catch (Exception e) { 
      System.out.println("Could not load XRayTool: " + e); 
      return null;
    } 
  }  // end of loadXRay()



  public static void invokeXRay(XScript xrayScript, Object obj) 
  {
    if (xrayScript == null) {
      System.out.println("XRayTool script is null");
      return;
    }

    try {
      // minimal inout/out parameters
      short[][] outParamIndex = { { 0 } };
      Object[][] outParam = { { null } };
      xrayScript.invoke(new Object[]{ obj }, outParamIndex, outParam);
    }
    catch (Exception e) { 
      System.out.println("Could not invoke XRayTool: " + e); 
    } 
  }  // end of invokeXRay()


  // -------------------- macro security --------------------


  public static int getSecurity()
  {
    System.out.println("Macro security level:");
    Integer val = (Integer) Info.getConfig(
                             "/org.openoffice.Office.Common/Security/Scripting",
                             "MacroSecurityLevel");
    if (val == null) {
      System.out.println("  Unknown");
      return 0;  // "low" security
    }
    else {
      int macroVal = val.intValue();

      if (macroVal == MacroExecMode.NEVER_EXECUTE)   
        // System.out.println("  Macros cannot be executed (0)");   
            // the documentation seems wrong; it doesn't match the "Macro Security" dialog
        System.out.println("  Macros executed without confirmation (0)");
 
      else if (macroVal == MacroExecMode.ALWAYS_EXECUTE)
        System.out.println("  Execute all macros; macros signed with trusted certificates or from the secure list are executed silently (2)");
      else if (macroVal == MacroExecMode.ALWAYS_EXECUTE_NO_WARN)
        System.out.println("  Silently execute all macros (4)");

      else if (macroVal == MacroExecMode.USE_CONFIG)
        System.out.println("  Use configuration for macro settings (3)");
      else if (macroVal == MacroExecMode.USE_CONFIG_REJECT_CONFIRMATION)
        System.out.println("  Treat cases where confirmation required as macro rejection (5)");
      else if (macroVal == MacroExecMode.USE_CONFIG_APPROVE_CONFIRMATION)
        System.out.println("  Treat cases where confirmation required as macro approval (6)");
 
      else if (macroVal == MacroExecMode.FROM_LIST)
        System.out.println("  Execute macros in the installed list (1)");
      else if (macroVal == MacroExecMode.FROM_LIST_NO_WARN)
        System.out.println("  Silently execute macros in the installed list (7)");
 
      else if (macroVal == MacroExecMode.FROM_LIST_AND_SIGNED_WARN)
        System.out.println("  Execute macros in the secure list or macros signed by trusted certificates (8)");
      else if (macroVal == MacroExecMode.FROM_LIST_AND_SIGNED_NO_WARN)
        System.out.println("  Silently execute macros in the secure list or macros signed by trusted certificates (9)");
      else
        System.out.println("  Unknown macro security value: " + macroVal);

      return macroVal;
    }
  }  // end of getSecurity()



  public static boolean setSecurity(int level)
  {
    if ((level == Macros.LOW) || (level == Macros.MEDIUM) || 
        (level == Macros.HIGH) || (level == Macros.VERY_HIGH)) {
      System.out.println("Setting macro security level to " + level);
      return Info.setConfig("/org.openoffice.Office.Common/Security/Scripting",
                             "MacroSecurityLevel", Integer.valueOf(level));
    }
    else {
      System.out.println("Use Macros class constants: LOW, MEDIUM, HIGH, or VERY_HIGH");
      return false;
    }
  }  // end of setSecurity()




  // ------------------- macros and events -------------------

  // ---- for office events -----

  public static void listOfficeEvents()
  { System.out.println("\nEvent Handler names");
    XNameReplace eventHandlers = getEventHandlers();
    Lo.printNames( eventHandlers.getElementNames() );
  }


  public static XNameReplace getEventHandlers()
  { XGlobalEventBroadcaster geb = theGlobalEventBroadcaster.get(Lo.getContext());
    return geb.getEvents(); 
  }


  public static PropertyValue[] getEventProps(String eventName)
  { 
    XNameReplace eventHandlers = getEventHandlers();
    return getEventProps( eventHandlers, eventName);
  } 


  public static PropertyValue[] getEventProps(XNameReplace eventHandlers, String eventName)
  { 
    try {
      Object oProps = eventHandlers.getByName(eventName);
      if (AnyConverter.isVoid(oProps))  // needed or conversion will fail
        return null;
      else
        return (PropertyValue[])oProps;
    }
    catch(com.sun.star.uno.Exception e)
    {  System.out.println("Could not find event " + eventName);
       return null;
    }
  } // end of getEventProps()



  public static void setEventScript(String eventName, String scriptName)
  {
    PropertyValue[] evProps = getEventProps(eventName);
    if (evProps != null)
      Props.setProp(evProps, "Script", scriptName);
    else
      evProps = Props.makeProps("EventType", "Script",
                                "Script", scriptName);

    XNameReplace eventHandlers = getEventHandlers();
    try {
      eventHandlers.replaceByName(eventName, evProps);
      System.out.println("Set script for " + eventName + " to \"" +
                                             scriptName + "\"");
    }
    catch(com.sun.star.uno.Exception e)
    {  System.out.println("Could not set script for " + eventName);  }
  } // end of setEventScript()



  // ---- for document events ----

  public static void listDocEvents(Object odoc)
  { System.out.println("\nDoc Event Handler names");
    XNameReplace eventHandlers = getDocEventHandlers(odoc);
    Lo.printNames( eventHandlers.getElementNames() );
  }


  public static XNameReplace getDocEventHandlers(Object odoc)
  // was XComponent
  { XEventsSupplier es = Lo.qi(XEventsSupplier.class, odoc);
    return es.getEvents();
  }


  public static PropertyValue[] getDocEventProps(Object odoc, String eventName)
  { 
    XNameReplace eventHandlers = getDocEventHandlers(odoc);
    return getEventProps( eventHandlers, eventName);
 } 




  public static void setDocEventScript(Object odoc, String eventName, String scriptName)
  {
    PropertyValue[] evProps = getDocEventProps(odoc, eventName);
    if (evProps != null)
      Props.setProp(evProps, "Script", scriptName);
    else
      evProps = Props.makeProps("EventType", "Script",
                                "Script", scriptName);

    XNameReplace eventHandlers = getDocEventHandlers(odoc);
    try {
      eventHandlers.replaceByName(eventName, evProps);
      System.out.println("Set doc script for " + eventName + " to \"" +
                                                 scriptName + "\"");
    }
    catch(com.sun.star.uno.Exception e)
    {  System.out.println("Could not set doc script for " + eventName);  }
  } // end of setDocEventScript()


}  // end of Macros class


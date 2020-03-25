
// Dialogs.java
// Andrew Davison, October 2016, ad@fivedots.psu.ac.th


/* Provide methods to create, execute and dispose an OOo Dialog
   with labels (fixed text), buttons, text fields, and password fields.

   Mostly based on code from UnoDialogSample.java at
         http://api.libreoffice.org/examples/DevelopersGuide/examples.html#GraphicalUserInterfaces

   The method categories:
     * load & execute a dialog
     * access a control/component inside a dialog
     * convert dialog into other forms
     * create a dialog
     * add a component to a dialog: 
         label, button, text field, password field, 
         combo box, check box 

*/

package utils;

import com.sun.star.awt.*;
import com.sun.star.beans.*;
import com.sun.star.container.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;
import com.sun.star.frame.*;

import com.sun.star.uno.Exception;


public class Dialogs
{

  // -------------- load & execute a dialog -----------------


  public static XDialog loadDialog(String scriptName)
  // e.g. "Standard.MyHighlight"
  {
    XDialogProvider dp = Lo.createInstanceMCF(XDialogProvider.class, 
                                         "com.sun.star.awt.DialogProvider");
    if (dp == null) {
      System.out.println("Could not access the Dialog Provider");
      return null;
    }
    try {
      return dp.createDialog("vnd.sun.star.script:" + scriptName + 
                                             "?location=application");
    }
    catch (java.lang.Exception e) {
      System.out.println("Could not load the dialog: \"" + 
                                                scriptName + "\": " + e);
      return null;
    }
  }  // end of loadDialog()



  public static XDialog loadAddonDialog(String extensionID, String dialogFnm)
  // e.g. "org.openoffice.ezhighlightAddon", "dialogLibrary/foo.xdl"
  {
    XDialogProvider dp = Lo.createInstanceMCF(XDialogProvider.class, 
                                           "com.sun.star.awt.DialogProvider");
    if (dp == null) {
      System.out.println("Could not access the Dialog Provider");
      return null;
    }
    try {
      return dp.createDialog("vnd.sun.star.extension://" + 
                                          extensionID + "/" + dialogFnm);
    }
    catch (java.lang.Exception e) {
      System.out.println("Could not load the dialog: \"" + dialogFnm + "\": " + e);
      return null;
    }
  }  // end of loadAddonDialog()



  // --------- access a control/component inside a dialog -------------------


  public static XControl findControl(XControl dialogCtrl, String name)
  { XControlContainer ctrlCon = Lo.qi(XControlContainer.class, dialogCtrl);
    return ctrlCon.getControl(name);
  } 



  public static void showControlInfo(XControl dialogCtrl)
  {
    XControl[] controls = getDialogControlsArr(dialogCtrl);
    System.out.println("No of controls: " + controls.length);
    for (int i = 0; i < controls.length; i++) { 
      // Props.showProps("Properties for control " + i, getControlProps(controls[i]));
      System.out.println(i + ". Name: " + getControlName(controls[i]));
      System.out.println("  Default Control: " + getControlClassID(controls[i]));
      System.out.println();
    }
  }  // end of showControlInfo()



  public static XControl[] getDialogControlsArr(XControl dialogCtrl)
  { 
    XControlContainer ctrlCon = Lo.qi(XControlContainer.class, dialogCtrl);  
    return ctrlCon.getControls();
  }


  public static XPropertySet getControlProps(Object controlModel)
  {   return Lo.qi(XPropertySet.class, controlModel);   } 



  public static String getControlName(XControl control)
  {
    try {
      XPropertySet props = getControlProps(control.getModel());
      return (String) props.getPropertyValue("Name");
    }
    catch (com.sun.star.uno.Exception ex) {
      System.out.println("Could not access control's name");
      return null;
    }
  }  // end of getControlName()



  public static String getControlClassID(XControl control)
  {
    try {
      XPropertySet props = getControlProps(control.getModel());
      return (String) props.getPropertyValue("DefaultControl");
    }
    catch (com.sun.star.uno.Exception ex) {
      System.out.println("Could not access control's class ID");
      return null;
    }
  }  // end of getControlClassID()



  public static String getEventSourceName(EventObject event)
  {   return getControlName( getEventControl(event));  }


  public static XControl getEventControl(EventObject event)
  {  return Lo.qi(XControl.class, event.Source);  }


  // --------------- convert dialog into other forms ------------------


  public static XDialog getDialog(XControl dialogCtrl)
  {  return Lo.qi(XDialog.class, dialogCtrl);  }  


  public static XControl getDialogControl(XDialog dialog)
  {  return Lo.qi(XControl.class, dialog);  }  



  public static XTopWindow getDialogWindow(XControl dialogCtrl)
  // Lo.qi(XControlContainer.class, 
  {  return Lo.qi(XTopWindow.class, dialogCtrl);  }



  // ----------------- create a dialog ------------------------------


  public static XControl createDialogControl(int x, int y, int width, int height, String title)
  {
    try {
      XControl dialogCtrl =  Lo.createInstanceMCF(XControl.class, 
                                    "com.sun.star.awt.UnoControlDialog");
      XControlModel xControlModel = Lo.createInstanceMCF(XControlModel.class, 
                                     "com.sun.star.awt.UnoControlDialogModel");
      dialogCtrl.setModel(xControlModel);  // link view and model
      
      XPropertySet props = getControlProps(dialogCtrl.getModel());
          /* inherited from UnoControlDialogModel and its 
             super-superclass UnoControlDialogElement */
      props.setPropertyValue("PositionX", x);
      props.setPropertyValue("PositionY", y);
      props.setPropertyValue("Height", height);
      props.setPropertyValue("Width", width);

      props.setPropertyValue("Title", title);
      props.setPropertyValue("Name", "OfficeDialog");

      props.setPropertyValue("Step", 0);
      props.setPropertyValue("Moveable", true);
      props.setPropertyValue("TabIndex", new Short((short) 0));

      return dialogCtrl;
    }
    catch (Exception ex) {
      System.out.println("Could not create dialog control: " + ex);
      return null;
    }
  }  // end of createDialogControl()



  public static XDialog createDialogPeer(XControl dialogCtrl)
  // create a peer for this dialog, using the Office desktop window as a parent
  {
    XWindow xWindow = (XWindow) Lo.qi(XWindow.class, dialogCtrl);
    xWindow.setVisible(false);
                  // set the dialog window invisible until it is executed

    XToolkit xToolkit = Lo.createInstanceMCF(XToolkit.class, "com.sun.star.awt.Toolkit");
    XWindowPeer windowParentPeer = xToolkit.getDesktopWindow();

    // System.out.println("Dialogs.execute() toolkit:" + xToolkit);
    // System.out.println("Dialogs.execute() windowParentPeer:" + windowParentPeer);

    dialogCtrl.createPeer(xToolkit, windowParentPeer);
    // XWindowPeer peer = dialogCtrl.getPeer();

    XComponent dialogComponent = Lo.qi(XComponent.class, dialogCtrl);
    XDialog dialog = getDialog(dialogCtrl);
    // dialogComponent.dispose();      // free window resources
        /* commented out or the Add-on dialog crashes when called a second time
           because createPeer() cannot find a model */
    return dialog;
  }  // end of createDialogPeer()






  // ---------------- add components to a dialog -----------------------------
  // label, button, text field, password field, 
  // combo box, check box


  public static XControl insertLabel(XControl dialogCtrl,
                                           int x, int y, int width, String label)
  { try {
      XMultiServiceFactory msf = Lo.qi(XMultiServiceFactory.class, 
                                                       dialogCtrl.getModel());
      Object model = msf.createInstance("com.sun.star.awt.UnoControlFixedTextModel");

      XNameContainer nameCon = getDialogNmCon(dialogCtrl);
      String nm = createName(nameCon, "FixedText");
      
      // Set properties in the model
      XPropertySet props = getControlProps(model);
      props.setPropertyValue("PositionX", x);
      props.setPropertyValue("PositionY", y + 2);
      props.setPropertyValue("Height", 8);
      props.setPropertyValue("Width", width);
      props.setPropertyValue("Label", label);
      props.setPropertyValue("Name", nm);
      
      // Add the model to the dialog
      nameCon.insertByName(nm, model);
      
      // reference the control by name
      XControlContainer ctrlCon = Lo.qi(XControlContainer.class, dialogCtrl);
      return ctrlCon.getControl(nm);
    }
    catch (Exception ex) {
      System.out.println("Could not create fixed text control: " + ex);
      return null;
    }
  }  // end of insertLabel()



  public static XNameContainer getDialogNmCon(XControl dialogCtrl)
  {  return Lo.qi(XNameContainer.class, dialogCtrl.getModel());  }



  public static String createName(XNameAccess elemContainer, String name)
  // Make a unique string by appending a number to the supplied name
  {
    boolean usedName = true;
    int i = 1;
    String nm = name + i;
    while (usedName) {
      usedName = elemContainer.hasByName(nm);
      if (usedName) {
        i++;
        nm = name + i;
      }
    }
    return nm;
  }  // end of createName()



  public static XControl insertButton(XControl dialogCtrl,
                                         int x, int y, int width, String label)
  {  return insertButton(dialogCtrl, x, y, width, label, PushButtonType.STANDARD_value);  }


  public static XControl insertButton(XControl dialogCtrl,
                       int x, int y, int width, String label, int pushButtonType)
  { try {
      // create a button model
      XMultiServiceFactory msf = Lo.qi(XMultiServiceFactory.class, 
                                                       dialogCtrl.getModel());
      Object model = msf.createInstance("com.sun.star.awt.UnoControlButtonModel");
      
      // generate a unique name for the control
      XNameContainer nameCon = getDialogNmCon(dialogCtrl);
      String nm = createName(nameCon, "CommandButton");
      
      // set properties in the model
      XPropertySet props = getControlProps(model);
          // inherited from UnoControlDialogElement and UnoControlButtonModel
      props.setPropertyValue("PositionX", x);
      props.setPropertyValue("PositionY", y);
      props.setPropertyValue("Height", 14);
      props.setPropertyValue("Width", width);
      props.setPropertyValue("Label", label);
      props.setPropertyValue("PushButtonType", new Short((short) pushButtonType));
      props.setPropertyValue("Name", nm);
      
      // Add the model to the dialog
      nameCon.insertByName(nm, model);
      
      // get the dialog's container holding all the control views
      XControlContainer ctrlCon = Lo.qi(XControlContainer.class, dialogCtrl);

      // use the model's name to get its view inside the dialog
      return ctrlCon.getControl(nm);
    }
    catch (Exception ex) {
      System.out.println("Could not create button control: " + ex);
      return null;
    }
  }  // end of insertButton()



  public static XControl insertTextField(XControl dialogCtrl,
                                            int x, int y, int width, String text)
  {  return insertTextField(dialogCtrl, x, y, width, text, ' ');  }


  public static XControl insertPasswordField(XControl dialogCtrl,
                                            int x, int y, int width, String text)
  {  return insertTextField(dialogCtrl, x, y, width, text, '*'); }



  private static XControl insertTextField(XControl dialogCtrl, 
                                         int x, int y, int width, String text, char echoChar) 
  { try {
      XMultiServiceFactory msf = Lo.qi(XMultiServiceFactory.class, 
                                                       dialogCtrl.getModel());
      Object model = msf.createInstance("com.sun.star.awt.UnoControlEditModel");
      // System.out.println("text field model: " + model);

      XNameContainer nameCon = getDialogNmCon(dialogCtrl);
      String nm = createName(nameCon, "TextField");
      
      // Set the properties in the model
      XPropertySet props = getControlProps(model);
         // inherited from UnoControlDialogElement and UnoControlEditModel 
      props.setPropertyValue("PositionX", x);
      props.setPropertyValue("PositionY", y);
      props.setPropertyValue("Height", 12);
      props.setPropertyValue("Width", width);
      props.setPropertyValue("Text", text);
      props.setPropertyValue("Name", nm);
      
      if (echoChar == '*')    // for password fields
        props.setPropertyValue("EchoChar",  new Short((short) echoChar));
      
      // Add the model to the dialog
      nameCon.insertByName(nm, model);
      
      // reference the control by name
      XControlContainer ctrlCon = Lo.qi(XControlContainer.class, dialogCtrl);
      return ctrlCon.getControl(nm);
    }
    catch (Exception ex) {
      System.out.println("Could not create text field control: " + ex);
      return null;
    }
  }  // end of insertTextField()



  public static XControl insertComboBox(XControl dialogCtrl,
                               int x, int y, int width, String[] entries)
  { try {
      XMultiServiceFactory msf = Lo.qi(XMultiServiceFactory.class, 
                                                       dialogCtrl.getModel());
      Object model = msf.createInstance("com.sun.star.awt.UnoControlComboBoxModel");

      XNameContainer nameCon = getDialogNmCon(dialogCtrl);
      String nm = createName(nameCon, "ComboBox");

      XPropertySet props = getControlProps(model);
      props.setPropertyValue("PositionX", x);
      props.setPropertyValue("PositionY", y);
      props.setPropertyValue("Height", 12);
      props.setPropertyValue("Width", width);
      props.setPropertyValue("Name", nm);
      props.setPropertyValue("Dropdown", true);
      props.setPropertyValue("StringItemList", entries);
      props.setPropertyValue("MaxTextLen", new Short((short) 10));
      props.setPropertyValue("ReadOnly", false);

      // add the model to the dialog
      nameCon.insertByName(nm, model);

      // reference the control by name
      XControlContainer ctrlCon = Lo.qi(XControlContainer.class, dialogCtrl);
      return ctrlCon.getControl(nm);
    }
    catch (Exception ex) {
      System.out.println("Could not create combobox control: " + ex);
      return null;
    }
  }  // end of insertComboBox()




  public static XControl insertCheckBox(XControl dialogCtrl, 
                                     int x, int y, int width, String label)
  { try {
      XMultiServiceFactory msf = Lo.qi(XMultiServiceFactory.class, 
                                                       dialogCtrl.getModel());
      Object model = msf.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");

      XNameContainer nameCon = getDialogNmCon(dialogCtrl);
      String nm = createName(nameCon, "CheckBox");

      XPropertySet props = getControlProps(model);
      props.setPropertyValue("PositionX", x);
      props.setPropertyValue("PositionY", y);
      props.setPropertyValue("Height", 8);
      props.setPropertyValue("Width", width);
      props.setPropertyValue("Name", nm);
      props.setPropertyValue("Label", label);
      props.setPropertyValue("TriState", true);
      props.setPropertyValue("State", new Short((short) 1));

      // add the model to the dialog
      nameCon.insertByName(nm, model);

      // reference the control by name
      XControlContainer ctrlCon = Lo.qi(XControlContainer.class, dialogCtrl);
      return ctrlCon.getControl(nm);
    }
    catch (Exception ex) {
      System.out.println("Could not create check box control: " + ex);
      return null;
    }
  }  // end of insertCheckBox()


}  // end of Dialogs class

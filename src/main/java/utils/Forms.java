
// Forms.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2016

/* Functions for:
     - accessing forms in document
     - get Form models
     - get the control for a model
     - creating controls
     - bind form to database
     - bind a macro to a form control
*/

package utils;

import java.util.ArrayList;

import com.sun.star.beans.*;
import com.sun.star.container.*;
import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;
import com.sun.star.util.*;
import com.sun.star.drawing.*;
import com.sun.star.awt.*;
import com.sun.star.text.*;
import com.sun.star.view.*;

import com.sun.star.uno.Exception;
// import com.sun.star.io.IOException;

import com.sun.star.sdb.*;

import com.sun.star.form.*;
import com.sun.star.form.XLoadable;

import com.sun.star.script.*;


public class Forms
{
  // ------------------------ access forms in document -----------------------


  public static XNameContainer getForms(XComponent doc)
  {
    XDrawPage drawPage = getDrawPage(doc);
    if (drawPage != null)
      return getForms(drawPage);
    else
      return null;
  }  // end of getForms()




  public static XDrawPage getDrawPage(XComponent doc)
  // return the first draw page even if there are many
  {
    XDrawPageSupplier xSuppPage = Lo.qi(XDrawPageSupplier.class, doc);
                                // doc only supports a single DrawPage
    if (xSuppPage != null)
      return xSuppPage.getDrawPage();
    else {
      XDrawPagesSupplier xSuppPages = Lo.qi(XDrawPagesSupplier.class, doc);
                                 // doc supports multiple DrawPages
      XDrawPages xPages = xSuppPages.getDrawPages();
      try {
        // System.out.println("Returning first draw page of many pages");
        return Lo.qi(XDrawPage.class, xPages.getByIndex(0));
      }
      catch(Exception e) {
        System.out.println(e);
        return null;
      }
    }
  }  // end of getDrawPage()



  public static XNameContainer getForms(XDrawPage drawPage)
  // get all the forms in the page as a named container
  {
    XFormsSupplier formsSupp = Lo.qi(XFormsSupplier.class, drawPage);
    return formsSupp.getForms();
  } // end of getForms()




  public static XNameContainer getForm(XComponent doc)
  // get the *first form* in the page as as a named container
  {
    XDrawPage drawPage = getDrawPage(doc);
    if (drawPage != null)
      return getForm(drawPage);
    else
      return null;
  }  // end of getForm()




  public static XNameContainer getForm(XDrawPage drawPage)
  // get the *first form* in the page as as a named container
  {
    if (drawPage == null) {
      System.out.println("No draw page supplied");
      return null;
    }
    XIndexContainer idxForms = getIndexedForms(drawPage);
    if (idxForms == null) {
      System.out.println("No forms found in draw page");
      return null;
    }
    try {
      return Lo.qi(XNameContainer.class, idxForms.getByIndex(0));
    }
    catch(Exception e){
      System.out.println("Could not find default form");
      return null;
    }
  } // end of getForm()




  public static XForm getForm(XComponent doc, String formName)
  // get the form called formName
  {
     XNameContainer namedForms = Forms.getForms(doc);
     XNameContainer con = Forms.getFormByName(formName, namedForms);
     return Lo.qi(XForm.class, con);
  }



  public static XNameContainer getFormByName(String formName,
                                         XNameContainer namedForms)
  // get the form called formName
  {
    try {
      return Lo.qi(XNameContainer.class, namedForms.getByName(formName));
    }
    catch (Exception e) {
      System.out.println("Could not find the form " + formName + ": " + e);
      return null;
    }
  }  // end of getFormByName()



  private static XIndexContainer getIndexedForms(XDrawPage drawPage)
  {
    XFormsSupplier formsSupp =
              Lo.qi(XFormsSupplier.class, drawPage);
    return Lo.qi(XIndexContainer.class, formsSupp.getForms());
  } // end of getIndexedForms()



  public static XNameContainer insertForm(String formName, XComponent doc)
  // add a data form called formName to the document
  {
    XNameContainer docForms = Forms.getForms(doc);
    return insertForm("GridForm", docForms);
  }


  public static XNameContainer insertForm(String formName,
                                          XNameContainer namedForms)
  // add a data form called formName to the forms
  { try {
      if (!namedForms.hasByName(formName)) {
        XNameContainer xNamedForm = Lo.createInstanceMSF(XNameContainer.class,
                                        "com.sun.star.form.component.DataForm");
        namedForms.insertByName(formName, xNamedForm);
        return xNamedForm;
      }
      else {
        System.out.println(formName + " already exists");
        return getFormByName(formName, namedForms);
      }
    }
    catch (Exception e) {
      System.out.println("Could not insert the form " + formName + ": " + e);
      return null;
    }
  }  // end of insertForm()





  public static boolean hasForm(XComponent doc, String formName)
  // is the formName form in the doc?
  {
    XDrawPage drawPage = getDrawPage(doc);
    if (drawPage == null)
      return false;
    XNameContainer namedFormsContainer = getForms(drawPage);
    if(namedFormsContainer == null) {
      System.out.println("No forms found on page");
      return false;
    }
    XNameAccess xNamedForms =
          Lo.qi(XNameAccess.class, namedFormsContainer);
    return xNamedForms.hasByName(formName);
  }  // end of hasForm()




  public static void showFormNames(XComponent doc)
  {
    XNameContainer formNamesCon = Forms.getForms(doc);
    String[] formNames = formNamesCon.getElementNames();
    System.out.println("No. of forms found: " + formNames.length);
    for(String formName : formNames)
      System.out.println("  " + formName);
  }  // end of showFormNames()



  public static void listForms(XComponent doc)
  { XNameContainer formNamesCon = Forms.getForms(doc);
    listForms(formNamesCon, "  ");  
    System.out.println();
  }


  public static void listForms(XNameAccess xContainer, String tabStr)
  {
    String nms[] = xContainer.getElementNames();
    for (int i=0; i < nms.length; i++) {
      try {
        XServiceInfo servInfo = Lo.qi(XServiceInfo.class,
                                        xContainer.getByName(nms[i]));
        if (servInfo.supportsService("com.sun.star.form.FormComponents")) {
          // this means that a form has been found
          if (Info.supportService(servInfo, "com.sun.star.form.component.DataForm"))
            System.out.println(tabStr + "Data Form \"" + nms[i] + "\"");
          else
            System.out.println(tabStr + "Form \"" + nms[i] + "\"");
          // Info.showServices("Form", servInfo);
          // Info.showInterfaces("Form", servInfo);
          XNameAccess childCon = Lo.qi(XNameAccess.class, servInfo);
          listForms(childCon, tabStr + "  ");  // recursively list form components
        }
        else if (servInfo.supportsService("com.sun.star.form.FormComponent")) {
          XControlModel model = Lo.qi(XControlModel.class, servInfo);
          System.out.println(tabStr + "\"" + nms[i] + "\": " +
                                          Forms.getTypeStr(model));
          // Props.showObjProps("Model", model);
        }
        else
          System.out.println(tabStr + "unknown: " +  nms[i]);
      }
      catch(Exception e)
      {  System.out.println(tabStr + "Could not access " + nms[i]);  }
    }
  }  // end of listForms()




  // -----------------------------  get form models -----------------------



  public static ArrayList<XControlModel> getModels(XComponent doc)
  { XNameContainer formNamesCon = Forms.getForms(doc);
    return getModels(formNamesCon);
  }


  public static ArrayList<XControlModel> getModels(XNameAccess formNamesCon)
  // search recursive through the named forms container, collecting all
  // the control models into a list.
  {
    ArrayList<XControlModel> models = new ArrayList<XControlModel>();
    String nms[] = formNamesCon.getElementNames();
    for (int i=0; i < nms.length; i++) {
      try {
        XServiceInfo servInfo = Lo.qi(XServiceInfo.class,formNamesCon.getByName(nms[i]));

        if (servInfo.supportsService("com.sun.star.form.FormComponents")) {
          // this means that a form has been found
          XNameAccess childCon = Lo.qi(XNameAccess.class, servInfo);
          models.addAll( getModels(childCon) );  // recursively search 
        }
        else if (servInfo.supportsService("com.sun.star.form.FormComponent")) {
          XControlModel model = Lo.qi(XControlModel.class, servInfo);
          models.add(model);
        }
      }
      catch(Exception e)
      {  System.out.println("Could not access " + nms[i]);  }
    }
    return models;
  }  // end of getModels()





  public static ArrayList<XControlModel> getModels2(XComponent doc, String formName)
  // another way to obtain models, via the control shapes in the DrawPage
  {
    XDrawPage xDrawPage = getDrawPage(doc);
    if (xDrawPage == null) {
      System.out.println("No draw page found");
      return null;
    }

    ArrayList<XControlModel> models = new ArrayList<XControlModel>();
    try {
      for (int i = 0; i < xDrawPage.getCount(); i++) {
        XControlShape cShape = Lo.qi(XControlShape.class,
                                        xDrawPage.getByIndex(i));
        XControlModel model = cShape.getControl();
        if (belongsToForm(model, formName))
          models.add(model);
      }
    }
    catch (Exception e) {
      System.out.println("Could not collect control model: " + e);
    }

    int numModels = models.size();
    System.out.println("No. of control models found: " + numModels);
    if (numModels == 0)
      return null;
    else
      return models;
  }  // end of getModels2()



  public static String getEventSourceName(EventObject event)
  {
    XControl control = Lo.qi(XControl.class, event.Source);
    return getName(control.getModel());
  }



  public static XControlModel getEventControlModel(EventObject ev)
  {
    XControl xControl = Lo.qi(XControl.class, ev.Source);
    return xControl.getModel();
  }  // end of getEventControlModel()



  public static String getFormName(XControlModel cModel)
  {
    XChild xChild = Lo.qi(XChild.class, cModel);
    XNamed xNamed = Lo.qi(XNamed.class, xChild.getParent());
    return xNamed.getName();
  }


  public static boolean belongsToForm(XControlModel cModel, String formName)
  {  return getFormName(cModel).equals(formName);  } 





  public static String getName(XControlModel cModel)
  // returns the name of the given form component
  {  return (String) Props.getProperty(cModel, "Name");  }



  public static String getLabel(XControlModel cModel)
  {  return (String) Props.getProperty(cModel, "Label");  }



  public static String getTypeStr(XControlModel cModel)
  {
    int id = getID(cModel);
    if (id == -1)
      return null;

    XServiceInfo servInfo = Lo.qi(XServiceInfo.class, cModel);
    switch (id) {
      case FormComponentType.COMMANDBUTTON:
        return "Command button"; 
      case FormComponentType.RADIOBUTTON:
        return "Radio button"; 
      case FormComponentType.IMAGEBUTTON:
        return "Image button"; 
      case FormComponentType.CHECKBOX:
        return "Check Box"; 
      case FormComponentType.LISTBOX:
        return "List Box"; 
      case FormComponentType.COMBOBOX:
        return "Combo Box"; 
      case FormComponentType.GROUPBOX:
        return "Group Box"; 
      case FormComponentType.FIXEDTEXT:
        return "Fixed Text"; 
      case FormComponentType.GRIDCONTROL:
        return "Grid Control"; 
      case FormComponentType.FILECONTROL:
        return "File Control"; 
      case FormComponentType.HIDDENCONTROL:
        return "Hidden Control"; 
      case FormComponentType.IMAGECONTROL:
        return "Image Control"; 
      case FormComponentType.DATEFIELD:
        return "Date Field"; 
      case FormComponentType.TIMEFIELD:
        return "Time Field"; 
      case FormComponentType.NUMERICFIELD:
        return "Numeric Field"; 
      case FormComponentType.CURRENCYFIELD:
        return "Currency Field"; 
      case FormComponentType.PATTERNFIELD:
        return "Pattern Field"; 
      case FormComponentType.TEXTFIELD:
        // two services with this class id: text field and formatted field 
        if ((servInfo != null) && servInfo.supportsService(
                       "com.sun.star.form.component.FormattedField"))
          return "Formatted Field";
        else
          return "Text Field";
      default: 
        System.out.println("Unknown class ID: " + id);
        return null;
    }
  }  // end of getTypeStr()



  public static int getID(XControlModel cModel)
  {
    // get the ClassId property
    Short classId = (Short) Props.getProperty(cModel, "ClassId");
    if (classId == null) {
      System.out.println("No class ID found for form component");
      return -1;
    } 
    return classId.intValue();
  }  // end of getID()




  public static boolean isButton(XControlModel cModel)
  {
    int id = getID(cModel);
    if (id == -1)
      return false;
    else 
      return ((id == FormComponentType.COMMANDBUTTON) ||
              (id == FormComponentType.IMAGEBUTTON));
  }  // end of isButton()



  public static boolean isTextField(XControlModel cModel)
  {
    int id = getID(cModel);
    if (id == -1)
      return false;
    else 
      return ((id == FormComponentType.DATEFIELD) ||
              (id == FormComponentType.TIMEFIELD) ||
              (id == FormComponentType.NUMERICFIELD) ||
              (id == FormComponentType.CURRENCYFIELD) ||
              (id == FormComponentType.PATTERNFIELD) ||
              (id == FormComponentType.TEXTFIELD));
  }  // end of isTextField()
      


  public static boolean isBox(XControlModel cModel)
  {
    int id = getID(cModel);
    if (id == -1)
      return false;
    else 
      return ((id == FormComponentType.RADIOBUTTON) ||
              (id == FormComponentType.CHECKBOX));
  }  // end of isBox()



  public static boolean isList(XControlModel cModel)
  {
    int id = getID(cModel);
    if (id == -1)
      return false;
    else 
      return ((id == FormComponentType.LISTBOX) ||
              (id == FormComponentType.COMBOBOX));
  }  // end of isList()



  /*  Other control types
      FormComponentType.GROUPBOX
      FormComponentType.FIXEDTEXT
      FormComponentType.GRIDCONTROL
      FormComponentType.FILECONTROL
      FormComponentType.HIDDENCONTROL
      FormComponentType.IMAGECONTROL
      FormComponentType.SCROLLBAR 
      FormComponentType.SPINBUTTON  
      FormComponentType.NAVIGATIONBAR
  */



  // --------------------- get control for a model --------------------



  public static XControl getControl(XComponent doc, XControlModel cModel)
  {
    XControlAccess controlAccess = GUI.getControlAccess(doc);
    if (controlAccess == null) {
      System.out.println("Could not obtain controls access in document");
      return null;
    }

    try {
      return controlAccess.getControl(cModel);
    }
    catch (Exception e) {
      System.out.println("Could not access control: " + e);
      return null;
    }
  }  // end of getControl()



  public static XControl getNamedControl(XComponent doc, String ctrlName)
  {
    ArrayList<XControlModel> models = Forms.getModels(doc);
    for(XControlModel model : models) {
      if (getName(model).equals(ctrlName)) {
        System.out.println("Found: " + ctrlName);
        return getControl(doc, model);
      }
    }
    System.out.println("No control found called " + ctrlName);
    return null;
  }  // end of getNamedControl()



  public static XControlModel getControlModel(XComponent doc, String ctrlName)
  {
    XControl control = getNamedControl(doc, ctrlName);
    if (control == null)
      return null;
    else
      return control.getModel();
  }  // end of getControlModel()


  // ----------------- create controls --------------------



  public static XPropertySet addControl(XComponent doc,  
                  String name, String label, String compKind, int x, int y,
                  int width, int height, XNameContainer parentForm)
  /* 
     add a control (view) and model (component) to the drawpage of the doc
       - create and initialize a control shape for the control view
       - create a control model
       - link the model to the shape
       - insert the shape into the shapes collection of the doc's draw page

     compKind: the service name of the control model, e.g. "TextField".

     parentForm: the form to use as the parent (container) for the control.
     If it's null, then the default form ("Form") will be used by default.

     Returns the property set for the control's model
  */
  {
    XPropertySet modelProps = null;
    try {
      // create a shape to represent the control's view
      XControlShape cShape =
          Lo.createInstanceMSF(XControlShape.class, "com.sun.star.drawing.ControlShape");

      // position and size of the shape
      cShape.setSize(new Size(width*100, height * 100));
      cShape.setPosition(new Point(x * 100, y * 100));

      // adjust the anchor so that the control is tied to the page
      XPropertySet shapeProps = Lo.qi(XPropertySet.class, cShape);
      TextContentAnchorType eAnchorType = TextContentAnchorType.AT_PARAGRAPH;

      shapeProps.setPropertyValue("AnchorType", eAnchorType);

      // create the control's model
      XControlModel cModel = Lo.createInstanceMSF(XControlModel.class, 
                                          "com.sun.star.form.component." + compKind);

      // insert the model into the form (or default to "Form")
      if (parentForm != null)
        parentForm.insertByName(name, cModel);

      // link model to the shape
      cShape.setControl(cModel);

      // add the shape to the shapes on the doc's draw page
      XDrawPage drawPage = getDrawPage(doc);
      XShapes formShapes = Lo.qi(XShapes.class, drawPage);
      formShapes.add(cShape);

      // set Name and Label properties for the model
      modelProps = Lo.qi(XPropertySet.class, cModel);
      modelProps.setPropertyValue("Name", name);
      if (label != null)
        modelProps.setPropertyValue("Label", label);
    }
    catch (Exception e) {
      System.out.println(e);
    }
    return modelProps;
  }  // end of addControl()



  public static XPropertySet addControl(XComponent doc, 
                  String name, String label, String compKind, int x, int y,
                                                     int width, int height)
  // use the default form, "Form", for the control
  {  return addControl(doc, name, label, compKind, x, y, width, height, null);  }





  public static XPropertySet addLabelledControl(XComponent doc, 
                    String label, String compKind, int x, int y, int height)
  // create a label and data field control, with the label preceding the control
  {
    XPropertySet ctrlProps = null;
    try {
      // create label (fixed text) control
      String name = label + "_Label";
      XPropertySet labelProps = addControl(doc, name, label, "FixedText", x, y, 25, 6);

      // create data field control
      ctrlProps = addControl(doc, label, null, compKind, x + 26, y, 40, height);
      ctrlProps.setPropertyValue("DataField", label);

      // add label props to the control
      ctrlProps.setPropertyValue("LabelControl", labelProps);
    }
    catch (Exception e) {
      System.out.println(e);
    }

    return ctrlProps;
  }  // end of addLabelledControl()




  public static XPropertySet addLabelledControl(XComponent doc, String label, String compKind, 
                                                     int y)
  {  return addLabelledControl(doc, label, compKind, 2, y, 6);  }



  public static XPropertySet addButton(XComponent doc, 
                                  String name, String label, int x, int y, int width)
  {  return addButton(doc, name, label, x, y, width, 6);  }


  public static XPropertySet addButton(XComponent doc, 
                      String name, String label, int x, int y, int width, int height)
  {
    XPropertySet buttonProps = null;
    try {
      buttonProps = addControl(doc, name, label, "CommandButton", x, y, width, height);
      buttonProps.setPropertyValue("HelpText", name);

      // don't want button to be accessible by the "tab" key
      buttonProps.setPropertyValue("Tabstop", false);

      // the button should not steal focus when clicked
      buttonProps.setPropertyValue("FocusOnClick", false);
    }
    catch (Exception e) {
      System.out.println(e);
    }

    return buttonProps;
  }  // end of addButton()




  public static XPropertySet addList(XComponent doc, 
                                  String name, String[] entries,
                                  int x, int y, int width, int height)
  // a list using a string array as its data source
  {
    XPropertySet listProps = null;
    try {
      listProps = addControl(doc, name, null, "ListBox", x, y, width, height);
      listProps.setPropertyValue("DefaultSelection", new short[]{0});
      listProps.setPropertyValue("ListSource", entries);
      listProps.setPropertyValue("Dropdown", true);
      listProps.setPropertyValue("MultiSelection", false);

      listProps.setPropertyValue("StringItemList", entries);
      listProps.setPropertyValue("SelectedItems", new short[]{0});
    }
    catch (Exception e) {
      System.out.println(e);
    }
    return listProps;
  }  // end of addList()




  public static XPropertySet addDatabaseList(XComponent doc, 
                                  String name, String sqlCmd,
                                  int x, int y, int width, int height)
  // a list using an SQL command as its data source
  {
    XPropertySet listProps = null;
    try {
      listProps = addControl(doc, name, null, "DatabaseListBox", x, y, width, height);
      // listProps.setPropertyValue("DefaultSelection", new short[]{0});  // causes hang
      listProps.setPropertyValue("Dropdown", true);
      listProps.setPropertyValue("MultiSelection", false);
      listProps.setPropertyValue("BoundColumn", (short) 0);

      // data-aware properties
      listProps.setPropertyValue("ListSourceType", ListSourceType.SQL);
      listProps.setPropertyValue("ListSource", new String[] { sqlCmd });
    }
    catch (Exception e) {
      System.out.println(e);
    }
    return listProps;
  }  // end of addDatabaseList()



  public static void createGridColumn(XControlModel gridModel, String dataField, 
                                                          String colKind, int width)
  /* adds a column to a grid;
     dataField: the database field to which the column should be bound;
     colKind: the column type (e.g. "NumericField");
     width: the column width (in mm). If 0, no width is set.
  */
  { 
    try {
      // column container and factory
      XIndexContainer colContainer = Lo.qi(XIndexContainer.class, gridModel);
      XGridColumnFactory colFactory = Lo.qi(XGridColumnFactory.class, gridModel);
      
      // create the column
      XPropertySet colProps = colFactory.createColumn(colKind);
      colProps.setPropertyValue("DataField", dataField);  
                                     // the field the column is bound to
      colProps.setPropertyValue("Label", dataField);
      colProps.setPropertyValue("Name", dataField);
      if (width > 0)
        colProps.setPropertyValue("Width", new Integer(width * 10));
      
      // add properties column to container
      colContainer.insertByIndex(colContainer.getCount(), colProps);
    }
    catch (Exception e) {
      System.out.println(e);
    }
  }  // end of createGridColumn()





  // -------------------- bind form to database ------------------



  public static void bindFormToTable(XForm xForm,
                                 String sourceName, String tableName)
  // Bind the form to the database in the sourceName URL
  {
    Props.setProperty(xForm, "DataSourceName", sourceName);
    Props.setProperty(xForm, "Command", tableName); // any table name
    Props.setProperty(xForm, "CommandType", CommandType.TABLE);
  }  // end of bindFormToTable()



  public static void bindFormToSQL(XForm xForm,
                                   String sourceName, String cmd)
  // Bind the form to the database in the sourceName URL, and send a SQL cmd
  {
    Props.setProperty(xForm, "DataSourceName", sourceName);
    Props.setProperty(xForm, "Command", cmd);   // SQL statement
    Props.setProperty(xForm, "CommandType", CommandType.COMMAND);  
                    // cannot use CommandType.TABLE for the SELECT cmd
  }  // end of bindFormToSQL()



  // -------------------- bind a macro to a form control ------------------


  public static void assignScript(XPropertySet controlProps, 
                              String interfaceName, String methodName, 
                              String scriptName, String loc)
  /* loc can be user, share, document, and extensions;
     see "Scripting Framework URI Specification"
     https://wiki.openoffice.org/wiki/Documentation/DevGuide/Scripting/
                                          Scripting_Framework_URI_Specification
  */
  {
    try {
      XChild propsChild = Lo.qi(XChild.class, controlProps);
      XIndexContainer parentForm = Lo.qi(XIndexContainer.class, propsChild.getParent());

      int pos = -1;
      for (int i = 0; i < parentForm.getCount(); i++) {
        XPropertySet child = Lo.qi(XPropertySet.class, parentForm.getByIndex(i) );
        if (UnoRuntime.areSame(child, controlProps)) {
          pos = i;
          break;
        }
      }
   
      if (pos == -1)
        System.out.println("Could not find contol's position in form");
      else {
        XEventAttacherManager manager = Lo.qi(XEventAttacherManager.class, parentForm);
        manager.registerScriptEvent(pos, 
             new ScriptEventDescriptor(interfaceName, methodName,
                                       "", "Script",
                     "vnd.sun.star.script:"+scriptName + 
                         "?language=Java&location=" + loc));
      }
    }
    catch( com.sun.star.uno.Exception e ) 
    { System.out.println(e);  }
  }  // end of assignScript()



}  // end of Forms class

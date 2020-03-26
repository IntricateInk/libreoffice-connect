
// Chart.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2014

/* A growing collection of utility functions to make Office
   easier to use. They are currently divided into the following
   groups:

     * chart utils

*/

package utils;

import java.io.*;
import java.awt.Point;
import java.util.*;
import java.awt.image.*;
import javax.imageio.*;
import java.util.regex.*;

import com.sun.star.beans.*;
import com.sun.star.comp.helper.*;
import com.sun.star.frame.*;
import com.sun.star.bridge.*;
import com.sun.star.lang.*;
import com.sun.star.text.*;
import com.sun.star.uno.*;
import com.sun.star.awt.*;
import com.sun.star.util.*;
import com.sun.star.drawing.*;
import com.sun.star.document.*;
import com.sun.star.container.*;
import com.sun.star.linguistic2.*;
import com.sun.star.graphic.*;
import com.sun.star.sheet.*;
import com.sun.star.style.*;
import com.sun.star.table.*;
import com.sun.star.embed.*;

import com.sun.star.chart.*;     // using chart 
// import com.sun.star.chart2.*;

import com.sun.star.uno.Exception;
import com.sun.star.io.IOException;


public class Chart
{

  // private static final String CHART_CLASSID = "12dcae26-281f-416f-a234-c3086127382e";

/*
  // chart label type
  public static final int NO_LABEL = 0;
  public static final int SHOW_NUMBER = 1;
  public static final int SHOW_PERCENT = 2;
  public static final int SHOW_CATEGORY = 4;
  public static final int CHECK_LEGEND = 16;
*/

  public static XChartDocument insertChart(XSpreadsheet sheet, String chartName,
                                           CellRangeAddress cellsRange, 
                                           int width, int height, String diagramName)
  {  return insertChart(sheet, chartName, cellsRange, 1, 1, width, height, diagramName);  }




  public static XChartDocument insertChart(XSpreadsheet sheet, String chartName,
                                           CellRangeAddress cellsRange, int x, int y, 
                                           int width, int height, String diagramName)
  /* Insert a chart using the given name as name of the OLE object and
    the range as corresponding
    range of data to be used for rendering.  The chart is placed in the sheet
    for charts at position (1,1) extending as large as given in chartSize.

    The diagram name must be the name of a diagram service (i.e one
    in "com.sun.star.chart.") that can be
    instantiated via the factory of the chart document
  */
  { 
    XTableCharts tableCharts = getTableCharts(sheet);
    XNameAccess tcAccess = Lo.qi(XNameAccess.class, tableCharts);
    if (tcAccess == null) {
      System.out.println("Unable to get name access to chart table");
      return null;
    }

    if (tcAccess.hasByName(chartName)) {
      System.out.println("A chart table called " + chartName + " already exists");
      return null;
    }

    // the unit for measures is 1/100th of a millimeter
    Rectangle rect = new Rectangle(x*1000, y*1000, width*1000, height*1000);

    CellRangeAddress[] addrs = new CellRangeAddress[]{ cellsRange };

    // first boolean: has column headers?; second boolean: has row headers?
    tableCharts.addNewByName(chartName, rect, addrs, true, true);
            /* 2nd last arg: whether the topmost row of the source data will be used 
               to set labels for the category axis or the legend.
               last arg: whether the leftmost column of the source data will be 
               used to set labels for the category axis or the legend.
            */

    XTableChart tableChart = getTableChart(tcAccess, chartName);
    XChartDocument chartDoc = getChartDoc(tableChart);
    if (chartDoc == null)
      return null;

    if (diagramName != null)
      setChartType(chartDoc, diagramName);

    return chartDoc;
  }  // end of insertChart()



  public static XNameAccess getTableChartAccess(XSpreadsheet sheet)
  {
    XTableCharts tableCharts = getTableCharts(sheet);
    return Lo.qi(XNameAccess.class, tableCharts);
  } // end of getTableChartAccess()



  public static XTableCharts getTableCharts(XSpreadsheet sheet)
  {
    // get the supplier for the charts
    XTableChartsSupplier chartsSupplier = Lo.qi(
                     XTableChartsSupplier.class, sheet);
    return chartsSupplier.getCharts();
  }


  public static String[] getChartNameList(XSpreadsheet sheet)
  {
    XTableCharts tableCharts = getTableCharts(sheet);
    if (tableCharts == null)
      return null;
    else
      return tableCharts.getElementNames();
  }  // end of getChartNameList()




  public static XTableChart getTableChart(XNameAccess tcAccess, String chartName)
  {
    XTableChart tableChart = null;
    try {
      tableChart = Lo.qi(XTableChart.class, tcAccess.getByName(chartName));
      // System.out.println("Found a table chart called " + chartName);
    }
    catch(Exception ex) 
    {  System.out.println("Could not access " + chartName); }
    return tableChart;
  }  // end of getTableChart()
  


  public static XChartDocument getChartDoc(XTableChart tableChart)
  {
    // the table chart is an embedded object which contains the chart document
    XEmbeddedObjectSupplier eos = Lo.qi(
                                         XEmbeddedObjectSupplier.class, tableChart);
    XInterface intf = eos.getEmbeddedObject();
    return Lo.qi(XChartDocument.class, intf);
  }  // end of getChartDoc()





  public static void setChartType(XChartDocument chartDoc, String diagramName)
  /* diagramName can be:
        BarDiagram, AreaDiagram, LineDiagram
        PieDiagram, DonutDiagram, NetDiagram
        XYDiagram, StockDiagram, BubbleDiagram, FilledNetDiagram
  */
  { try {
      XMultiServiceFactory msf = Lo.qi(
                                        XMultiServiceFactory.class, chartDoc);
      String chartType = "com.sun.star.chart." + diagramName;
      XDiagram diagram = Lo.qi(
                           XDiagram.class, msf.createInstance(chartType));
      chartDoc.setDiagram(diagram);
    }
    catch(Exception ex) {
      System.out.println("Could not set the chart type to " + diagramName);
    }
  }  // end of setChartType()



  public static String getChartType(XChartDocument chartDoc)
  {   return chartDoc.getDiagram().getDiagramType();  } 




  public static XChartDocument getChartDoc(XSpreadsheet sheet, String chartName)
  {
    XNameAccess tcAccess = getTableChartAccess(sheet);
    if (tcAccess == null) {
      System.out.println("Unable to get name access to table chart");
      return null;
    }

    if (!tcAccess.hasByName(chartName)) {
      System.out.println("No table chart called " + chartName + " found");
      return null;
    }

    XTableChart tableChart = getTableChart(tcAccess, chartName);
    if (tableChart == null)
      return null;
    else
      return getChartDoc(tableChart);
  }  // end of getChartDoc()



  public static boolean removeChart(XSpreadsheet sheet, String chartName)
  { 
    XTableCharts tableCharts = getTableCharts(sheet);
    XNameAccess tcAccess = Lo.qi(XNameAccess.class, tableCharts);
    if (tcAccess == null) {
      System.out.println("Unable to get name access to chart table");
      return false;
    }

    if (tcAccess.hasByName(chartName)) {
      tableCharts.removeByName(chartName);
      System.out.println("Chart table " + chartName + " removed");
      return true;
    }
    else {
      System.out.println("Chart table " + chartName + " not found");
      return false;
    }
  }  // end of removeChart()




  public static void setVisible(XSpreadsheet sheet, boolean isVisible)
  {
    //get draw page supplier for chart sheet 
    XDrawPageSupplier pageSupplier =  
           Lo.qi(XDrawPageSupplier.class, sheet); 
    
    XDrawPage drawPage = pageSupplier.getDrawPage(); 
    int numShapes = drawPage.getCount(); 
    // System.out.println("No. of shapes: " + numShapes);
    XShape shape;
    String classID;
    for (int i=0; i < numShapes; i++) {
      try {
        shape = Lo.qi(XShape.class, drawPage.getByIndex(i)); 
        classID = (String) Props.getProperty(shape, "CLSID");
        // System.out.println("Class ID: \"" + classID + "\"");
        if (classID.toLowerCase().equals(Lo.CHART_CLSID)) {
          // System.out.println("Found a chart");
          Props.setProperty(shape, "Visible", isVisible);
        }
      }
      catch(Exception e) {}
    }
  }  // end of setVisible()



  public static XShape getChartShape(XSpreadsheet sheet)
  // return the first chart shape
  {
    //get draw page supplier for chart sheet 
    XDrawPageSupplier pageSupplier =  
           Lo.qi(XDrawPageSupplier.class, sheet); 
    
    XDrawPage drawPage = pageSupplier.getDrawPage(); 
    int numShapes = drawPage.getCount(); 
    // System.out.println("No. of shapes: " + numShapes);
    XShape shape = null;
    String classID;
    for (int i=0; i < numShapes; i++) {
      try {
        shape = Lo.qi(XShape.class, drawPage.getByIndex(i)); 
        classID = (String) Props.getProperty(shape, "CLSID");
        // System.out.println("Class ID: \"" + classID + "\"");
        if (classID.toLowerCase().equals(Lo.CHART_CLSID)) {
          // System.out.println("Found a chart");
          break;
        }
      }
      catch(Exception e) {}
    }
    if (shape != null)
      System.out.println("Found a chart");
    return shape;
  }  // end of getChartShape()




  // ----------------------- Chart inside Draw --------------------------


  public XChartDocument insertChart(XDrawPage slide,
                          int x, int y, int width, int height, String diagramName)
  {
    XShape aShape = Draw.addShape(slide, "OLE2Shape", x, y, width, height);
    if (aShape == null)
      return null;

    XChartDocument chartDoc = getChartDoc(aShape);
    if (chartDoc != null) // create a diagram
      setChartType(chartDoc, diagramName);

    return chartDoc;
  }  // end of insertChartInDraw()





  public static XChartDocument getChartDoc(XShape shape)
  {
    XChartDocument chartDoc = null;
    try {
      // change the OLE shape into a chart
      XPropertySet shapeProps = Lo.qi(XPropertySet.class, shape);
      if(shapeProps == null)  {
        System.out.println("Unable to access shape properties");
        return null;
      }

      // set the class id for charts
      // shapeProps.setPropertyValue("CLSID", Lo.CHART_CLSID);

      // retrieve the chart document as model of the OLE shape
      chartDoc = Lo.qi(XChartDocument.class,
                                    shapeProps.getPropertyValue("Model"));
    }
    catch(Exception ex)
    { System.out.println("Couldn't change the OLE shape into a chart: " + ex); }

    return chartDoc;
  }  // end of getChartDoc()


  // ------------------------------ add chart to text object --------------------




  public static XChartDocument insertChart(XTextDocument doc,
                       int x, int y, int width, int height, String diagramName)
  {
    XChartDocument chartDoc = null;
    try {
      // create TextContent for embedding the formula in the text
      XMultiServiceFactory msFactory = Lo.getServiceFactory();
      if (msFactory == null) {
        System.out.println("No service factory");
        return null;
      }

      Object embedObj = msFactory.createInstance("com.sun.star.text.TextEmbeddedObject");
      XTextContent textContent = Lo.qi( XTextContent.class, embedObj);
      if (textContent != null) {
        System.out.println("Could not create embedded text object");
        return null;
      }

      XPropertySet propsSet = Lo.qi(XPropertySet.class, textContent);
      propsSet.setPropertyValue("CLSID", Lo.CHART_CLSID);

      XText xText = doc.getText();
      XTextCursor xCursor = xText.createTextCursor();

      // insert embedded object in text -> object will be created
      xText.insertTextContent(xCursor, textContent, true);

      // set size and position
      XShape xShape = Lo.qi(XShape.class, textContent);
      xShape.setSize( new Size(width*1000, height*1000));

      propsSet.setPropertyValue("VertOrient", VertOrientation.NONE);
      propsSet.setPropertyValue("HoriOrient", HoriOrientation.NONE);
      propsSet.setPropertyValue("VertOrientPosition", y*1000);
      propsSet.setPropertyValue("HoriOrientPosition", x*1000);

      // retrieve the chart document as model of the OLE shape
      chartDoc = Lo.qi(XChartDocument.class,
                                              propsSet.getPropertyValue("Model"));
      if (chartDoc != null) // create a diagram
        setChartType(chartDoc, diagramName);
/*
      chartDoc.setDiagram(
               Lo.qi( XDiagram.class,
                      Lo.qi( XMultiServiceFactory.class,
                                       chartDoc)).createInstance(diagramName)));
*/
    }
    catch(Exception ex)
    { System.out.println("Could not insert chart into text document: " + ex); }

    return chartDoc;
  }  // end of addChart()



  // --------------------------- adjust properties ------------------------


  public static String getTitle(XChartDocument chartDoc)
  {  return (String) Props.getProperty(chartDoc.getTitle(), "String");  }


  public static void setTitle(XChartDocument chartDoc, String title)
  {  
    XShape titleShape = chartDoc.getTitle();
    Props.setProperty(titleShape, "String", title);  
    // Props.setProperty(titleShape, "HasMainTitle", true);
  } 


  public static String getSubTitle(XChartDocument chartDoc)
  {  return (String) Props.getProperty(chartDoc.getSubTitle(), "String");  }


  public static void setSubTitle(XChartDocument chartDoc, String subtitle)
  {
    XShape subtitleShape = chartDoc.getTitle();
    Props.setProperty(subtitleShape, "String", subtitle);  
    // Props.setProperty(subtitleShape, "HasSubTitle", true);
  } 


  public static void setXAxisTitle(XChartDocument chartDoc, String title)
  {
    XDiagram diagram = chartDoc.getDiagram();
    if (diagram == null) {
      System.out.println("No Chart diagram found");
      return;
    }
    Props.setProperty(diagram, "HasXAxisDescription", true);
    XAxisXSupplier xAxis = Lo.qi(XAxisXSupplier.class, diagram);
    XShape titleShape = xAxis.getXAxisTitle();
    Props.setProperty(titleShape, "String", title);  
    // Props.setProperty(titleShape, "HasXAxisTitle", true);
  }  // end of setXAxisTitle()



  public static void setYAxisTitle(XChartDocument chartDoc, String title)
  {
    XDiagram diagram = chartDoc.getDiagram();
    if (diagram == null) {
      System.out.println("No Chart diagram found");
      return;
    }
    Props.setProperty(diagram, "HasYAxisDescription", true);
    XAxisYSupplier yAxis = Lo.qi(XAxisYSupplier.class, diagram);
    XShape titleShape = yAxis.getYAxisTitle();
    Props.setProperty(titleShape, "String", title);  
    // Props.setProperty(titleShape, "HasYAxisTitle", true);
  }  // end of setYAxisTitle()



  public static void setSecondYAxis(XChartDocument chartDoc, String title)
  {
    XDiagram diagram = chartDoc.getDiagram();
    if (diagram == null) {
      System.out.println("No Chart diagram found");
      return;
    }
    Props.setProperty(diagram, "HasSecondaryYAxis", true);

    XTwoAxisYSupplier yAxis2 = Lo.qi(XTwoAxisYSupplier.class, diagram);

    XPropertySet y2Props = yAxis2.getSecondaryYAxis(); 
    Props.showProps("Second y-axis", y2Props);

    XPropertySet y2TitleProps = 
                Lo.qi(XPropertySet.class, yAxis2.getYAxisTitle()); 
    Props.setProperty(yAxis2.getYAxisTitle(), "String", title); 
    Props.showProps("Second y-axis title", y2TitleProps);
  }  // end of setSecondYAxis()

            


  public static void showLegend(XChartDocument chartDoc)
  {  Props.setProperty(chartDoc, "HasLegend", true);  }



  public static void useHorizontals(XChartDocument chartDoc)
  {
    XDiagram diagram = chartDoc.getDiagram();
    if (diagram == null) {
      System.out.println("No chart diagram found");
      return;
    }
    Props.setProperty(diagram, "Vertical", true);  
  }  // end of useHorizontals()


  public static void usesLines(XChartDocument chartDoc, boolean hasLines)
  {
    XDiagram diagram = chartDoc.getDiagram();
    if (diagram == null) {
      System.out.println("No chart diagram found");
      return;
    }
    Props.setProperty(diagram, "Lines", hasLines);  
  }  // end of usesLines()




  public static void setDataCaption(XChartDocument chartDoc, int labelTypes)
  {
    XDiagram diagram = chartDoc.getDiagram();
    if (diagram == null) {
      System.out.println("No chart diagram found");
      return;
    }
    Props.setProperty(diagram, "DataCaption", labelTypes);  
  }  // end of setDataCaption()



  public static void setDataPlacement(XChartDocument chartDoc, long placement)
  // doesn't work
  {
    XDiagram diagram = chartDoc.getDiagram();
    if (diagram == null) {
      System.out.println("No chart diagram found");
      return;
    }
    //XPropertySet propsSet = Lo.qi(XPropertySet.class, diagram);
    //Props.showProps("Diagram", propsSet);

    try {
      XChartDataArray dataArr = Lo.qi(
                                     XChartDataArray.class, chartDoc.getData());
      double aData[][] = dataArr.getData();
      int numPoints = aData.length;
      System.out.println("No of points: " + numPoints);

      for (int i=0; i < numPoints; i++) {
        // first parameter is the index of the point, the second one is the series
        XPropertySet pointProps = diagram.getDataPointProperties(i, 0);
                
        //pointProps.setPropertyValue("CharHeight", 14.0);
        //pointProps.setPropertyValue("CharWeight", FontWeight.BOLD);
        // pointProps.setPropertyValue("CharColor", 0x993366);
        pointProps.setPropertyValue("LabelPlacement", placement);
      }

      XPropertySet propsSet = (XPropertySet) diagram.getDataRowProperties(0);
                                       //     diagram.getDataPointProperties(0,0);  
                                                                     // for pie chart
      Props.setProperty(propsSet, "LabelPlacement", placement);
      // Props.setProperty(propsSet, "LabelSeparator", "\n");
      Props.showProps("Data Row", propsSet);
    } 
    catch (Exception e) {
       System.out.println("Could not get DataPointProperties: "+ e);
    } 

  }  // end of setDataPlacement()





  public static void printDataArray(XChartDocument chartDoc)
  {
    XChartData chartData = chartDoc.getData();
    XChartDataArray dataArr = Lo.qi(XChartDataArray.class, chartData);
    double data[][] = dataArr.getData();
    if (data == null)
      System.out.println("No data found");
    else {
      System.out.println("No. of Data columns: " + data[0].length);
      System.out.println("No. of Data rows: " + data.length);
      for(int row=0; row < data.length; row++) {
        for(int col=0; col < data[row].length; col++)
          System.out.print("  " + data[row][col]);
        System.out.println();
      }
    }

    String[] rowDescs = dataArr.getRowDescriptions();
    if (rowDescs == null)
      System.out.println("No row description found");
    else {
      System.out.println("No. of rows: " + rowDescs.length);
      for(String row : rowDescs)
        System.out.print("  \"" + row + "\"");
      System.out.println();
    }

    String[] colDescs = dataArr.getColumnDescriptions();
    if (colDescs == null)
      System.out.println("No column description found");
    else {
      System.out.println("No. of columns: " + colDescs.length);
      for(String col : colDescs)
        System.out.print("  \"" + col + "\"");
      System.out.println();
    }
  }  // end of printDataArray()



  public static void setAreaTransparency(XChartDocument chartDoc, int val)
  {
		// Set the transparency of chart areas surrounding the diagram
		// to 100% transparent.
    if ((val < 0) || (val > 100))
      System.out.println("Transparency range is 1-100");
    else
      Props.setProperty(chartDoc.getArea(), "FillTransparence", val);
  }



  public static void set3D(XChartDocument chartDoc, boolean is3D)
  {
    XDiagram diagram = chartDoc.getDiagram();
    if (diagram == null) {
      System.out.println("No chart diagram found");
      return;
    }
    Boolean isVertical = (Boolean) Props.getProperty(diagram, "Vertical");
    Props.setProperty(diagram, "Dim3D", is3D);
    Props.setProperty(diagram, "SolidType", ChartSolidType.CYLINDER);
                                   // also RECTANGULAR_SOLID, CONE, PYRAMID
    Props.setProperty(diagram, "Vertical", isVertical);    // reapply
  }  // end of set3D()




  public static void setSymbol(XChartDocument chartDoc, boolean hasSymbol)
  {
    XDiagram diagram = chartDoc.getDiagram();
    if (diagram == null) {
      System.out.println("No chart diagram found");
      return;
    }
    if (hasSymbol)
      Props.setProperty(diagram, "SymbolType", ChartSymbolType.AUTO);
    else // no symbol
      Props.setProperty(diagram, "SymbolType", ChartSymbolType.NONE);
  }  // end of setSymbol()



  public static void setTrend(XChartDocument chartDoc, ChartRegressionCurveType curveType)
  {
    XDiagram diagram = chartDoc.getDiagram();
    if (diagram == null) {
      System.out.println("No chart diagram found");
      return;
    }


  //  try {
      XChartDataArray dataArr = Lo.qi(
                                     XChartDataArray.class, chartDoc.getData());
      double aData[][] = dataArr.getData();
      int numPoints = aData.length;
      System.out.println("No of points: " + numPoints);

      // first parameter is the index of the point, the second one is the series
      // XPropertySet pointProps = diagram.getDataPointProperties(1, 1);
              
      // Props.showProps("Data Point 0", pointProps);


      //XPropertySet drPropsSet = (XPropertySet) diagram.getDataRowProperties(1);
      // Props.showObjProps("Data Row", drPropsSet);

/*
      XPropertySet regPropsSet =  Lo.qi(XPropertySet.class, 
                                Props.getProperty(drPropsSet, "DataRegressionProperties"));
      Props.setProperty(regPropsSet, "LineWidth", 200);
      Props.showObjProps("Data Regression", regPropsSet);
*/
      if ((curveType == ChartRegressionCurveType.NONE) ||
		  	  (curveType == ChartRegressionCurveType.LINEAR) || 
		  	  (curveType == ChartRegressionCurveType.LOGARITHM) || 
		  	  (curveType == ChartRegressionCurveType.EXPONENTIAL) || 
		  	  (curveType == ChartRegressionCurveType.POWER))
        Props.setProperty(diagram, "RegressionCurves", curveType);
      else 
        System.out.println("Did not recognize curve type: " + curveType);
 //   } 
 //   catch (Exception e) {
 //      System.out.println("Could not get Data Row Properties: "+ e);
 //   } 
  }  // end of setTrend()






  public static void setDataRowDescriptions(XChartDocument chartDoc, String[] descs)
  // works, but affects the graphic data point position
  {
    XChartData chartData = chartDoc.getData();
    XChartDataArray dataArr = Lo.qi(XChartDataArray.class, chartData);

    String[] rowDescs = dataArr.getRowDescriptions();
    if (rowDescs == null) {
      System.out.println("No row description found");
      return;
    }

    if (rowDescs.length != descs.length) {
      System.out.println("Row length mismatch; No. of rows == " + rowDescs.length);
      return;
    }
    dataArr.setRowDescriptions(descs);

    // chartDoc.attachData(chartData);
  }  // end of setDataRowDescriptions()



  public static void showDataRowProps(XChartDocument chartDoc)
  // useful for changing font size of data point labels
  {
    XDiagram diagram = chartDoc.getDiagram();
    if (diagram == null) {
      System.out.println("No chart diagram found");
      return;
    }
    //XPropertySet propsSet = Lo.qi(XPropertySet.class, diagram);
    //Props.showProps("Diagram", propsSet);

    try {
      XChartDataArray dataArr = Lo.qi(
                                     XChartDataArray.class, chartDoc.getData());
      double aData[][] = dataArr.getData();
      int numPoints = aData.length;
      System.out.println("No of points: " + numPoints);


      for (int i=0; i < numPoints; i++) {
        // first parameter is the index of the point, the second one is the series
        XPropertySet pointProps = diagram.getDataPointProperties(i, 0);
        //Props.showProps(i + ". Data Point", pointProps);
        //System.out.println("\n");

        //pointProps.setPropertyValue("CharHeight", 14.0);
        //pointProps.setPropertyValue("CharWeight", FontWeight.BOLD);
        // pointProps.setPropertyValue("CharColor", 0x993366);
        // pointProps.setPropertyValue("LabelPlacement", placement);
      }

      XPropertySet propsSet = (XPropertySet) diagram.getDataRowProperties(0);
                                       //     diagram.getDataPointProperties(0,0);  
                                                                     // for pie chart
      // Props.setProperty(propsSet, "LabelPlacement", placement);
      // Props.setProperty(propsSet, "LabelSeparator", "\n");
      Props.showProps("Data Row", propsSet);
      // propsSet.setPropertyValue("CharHeight", 14.0);
    } 
    catch (Exception e) {
       System.out.println("Could not get DataPointProperties: "+ e);
    } 
  }  // end of showDataRowProps()


}  // end of Chart class


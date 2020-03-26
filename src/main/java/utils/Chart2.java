
// Chart2.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, October 2015

/* Chart2 Utilities:
     * insert a chart
     * get a chart
     * titles and subtitles
     * axes
     * gridlines
     * legend
     * background colors
     * access data source and series
     * chart types
     * using a data source
     * using the data series point properties
     * regression
     * add data to a chart
     * chart shape and image
*/

package utils;

import java.awt.Point;
import java.util.*;
import java.awt.image.*;


import com.sun.star.uno.*;
import com.sun.star.beans.*;
import com.sun.star.comp.helper.*;
import com.sun.star.frame.XController;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.bridge.*;
import com.sun.star.lang.*;
import com.sun.star.text.*;

import com.sun.star.awt.*;
import com.sun.star.util.*;
import com.sun.star.drawing.*;
import com.sun.star.document.*;
import com.sun.star.container.*;
import com.sun.star.graphic.*;
import com.sun.star.sheet.*;
import com.sun.star.style.*;
import com.sun.star.table.*;
import com.sun.star.embed.*;
import com.sun.star.view.*;

import com.sun.star.chart2.*;     // using chart2 
import com.sun.star.chart2.data.*;
// import com.sun.star.chart.*;
import com.sun.star.chart.ErrorBarStyle;
import com.sun.star.chart.ChartDataRowSource;

import com.sun.star.uno.Exception;
import com.sun.star.io.IOException;
import utils.Calc;


public class Chart2
{
  public static final int X_AXIS = 0;
  public static final int Y_AXIS = 1;
  public static final int Z_AXIS = 2;


  // regression line types
  public static final int LINEAR = 0;
  public static final int LOGARITHMIC = 1;
  public static final int EXPONENTIAL = 2;
  public static final int POWER = 3;
  public static final int POLYNOMIAL = 4;
  public static final int MOVING_AVERAGE = 5;
     /* LINEAR, LOGARITHMIC, EXPONENTIAL, POWER,
        used for scaling types also */

  private static final int[] CURVE_KINDS = 
        { LINEAR, LOGARITHMIC, EXPONENTIAL, POWER, POLYNOMIAL, MOVING_AVERAGE };

  private static final String[] CURVE_NAMES = 
        { "Linear", "Logarithmic", "Exponential", "Power", "Polynomial", "Moving average" };


  // data point label types
  public static final int DP_NUMBER = 0;
  public static final int DP_PERCENT = 1;
  public static final int DP_CATEGORY = 2;
  public static final int DP_SYMBOL = 3;
  public static final int DP_NONE = 4;

  // data point label placement
  public static final int DP_ABOVE = 0;   // maybe 2
  public static final int DP_CENTER = 1;
  public static final int DP_LEFT = 4;
  public static final int DP_BELOW = 6;
  public static final int DP_RIGHT = 8;



  private static final String CHART_NAME = "chart$$_";

  private static final String[] LINE_STYLES = {
     "Ultrafine Dashed", "Fine Dashed", "2 Dots 3 Dashes",
     "Fine Dotted", "Line with Fine Dots", "3 Dashes 3 Dots",
     "Ultrafine Dotted", "Line Style 9", "2 Dots 1 Dash",
     "Dashed"  };



  // ----------------- insert a chart ----------------------------



  public static XChartDocument insertChart(XSpreadsheet sheet, 
                         CellRangeAddress cellsRange, String cellName, 
                         int width, int height, String diagramName)
  { 
    String chartName = CHART_NAME + (int)(Math.random()*10000);  
                       // random name

    addTableChart(sheet, chartName, cellsRange, cellName, width, height);

    // get newly created (empty) chart
    XChartDocument chartDoc = getChartDoc(sheet, chartName);

    // assign chart template to the chart's diagram
    // System.out.println("Using chart template: " + diagramName);
    XDiagram diagram = chartDoc.getFirstDiagram();
    XChartTypeTemplate ctTemplate = 
                       setTemplate(chartDoc, diagram, diagramName);
    if (ctTemplate == null)
      return null;

    boolean hasCats = hasCategories(diagramName);

    // initialize data source
    XDataProvider dp = chartDoc.getDataProvider();
    PropertyValue[] aProps = Props.makeProps(
          new String[] { "CellRangeRepresentation", "DataRowSource",
                         "FirstCellAsLabel" , "HasCategories" },
          new Object[] { Calc.getRangeStr(cellsRange, sheet),
                         ChartDataRowSource.COLUMNS, true, hasCats });
    XDataSource ds = dp.createDataSource(aProps);

    // add data source to chart template
    PropertyValue[] args = Props.makeProps("HasCategories", hasCats);
    ctTemplate.changeDiagramData(diagram, ds, args);

    // apply style settings to chart doc
    setBackgroundColors(chartDoc, Calc.PALE_BLUE, Calc.LIGHT_BLUE); 
                                  // background and wall colors

    if (hasCats)  // charts using x-axis categories 
      setDataPointLabels(chartDoc, Chart2.DP_NUMBER);  // show y-axis values

    // printChartTypes(chartDoc);
    return chartDoc;
  }  // end of insertChart()




  public static void addTableChart(XSpreadsheet sheet, String chartName,
                          CellRangeAddress cellsRange, 
                          String cellName, int width, int height)
  // create new table chart at a given cell name and size, using cellsRange
  {
    XTableChartsSupplier chartsSupplier = 
                              Lo.qi(XTableChartsSupplier.class, sheet);
    XTableCharts tableCharts = chartsSupplier.getCharts();

    com.sun.star.awt.Point pos = Calc.getCellPos(sheet, cellName);
    Rectangle rect = new Rectangle(pos.X, pos.Y, width*1000, height*1000);

    CellRangeAddress[] addrs = new CellRangeAddress[]{ cellsRange };
    tableCharts.addNewByName(chartName, rect, addrs, true, true);
     /* 2nd last arg: whether the topmost row of the source data will 
        be used to set labels for the category axis or the legend;
        last arg: whether the leftmost column of the source data will 
        be used to set labels for the category axis or the legend.
     */
  }  // end of addTableChart()




  public static XChartTypeTemplate setTemplate(
                XChartDocument chartDoc, 
                XDiagram diagram, String diagramName)
  // change diagram to use the specified chart template
  { try {
      XChartTypeManager ctMan = chartDoc.getChartTypeManager();
      XMultiServiceFactory msf = Lo.qi(XMultiServiceFactory.class, ctMan);
      String templateNm = "com.sun.star.chart2.template." + diagramName;    
      XChartTypeTemplate ctTemplate = 
            Lo.qi(XChartTypeTemplate.class, msf.createInstance(templateNm));
      if (ctTemplate == null) {
        System.out.println("Could not create chart template \"" + 
                   diagramName +  "\"; using a column chart instead");
        ctTemplate = Lo.qi(XChartTypeTemplate.class, 
               msf.createInstance("com.sun.star.chart2.template.Column"));
      }
      ctTemplate.changeDiagram(diagram);
      return ctTemplate;
    }
    catch(Exception ex) {
      System.out.println("Could not set the chart type to " + diagramName);
      return null;
    }
  }  // end of setTemplate()




  public static boolean hasCategories(String diagramName)
  /* All the chart templates, except for scatter and bubble use
     categories on the x-axis
  */
  {
    String name = diagramName.toLowerCase();
    if (name.contains("scatter") || name.contains("bubble"))
       return false;
    return true;
  }  // end of hasCategories()


  // ----------------- get a chart ----------------------------


  public static XChartDocument getChartDoc(XSpreadsheet sheet, 
                                                  String chartName)
  // return the chart doc from the sheet
  {
    // get the named table chart
    XTableChart tableChart = getTableChart(sheet, chartName);
    if (tableChart == null)
      return null;

    // chart doc is embedded inside table chart
    XEmbeddedObjectSupplier eos = 
                Lo.qi(XEmbeddedObjectSupplier.class, tableChart);
    return Lo.qi(XChartDocument.class, eos.getEmbeddedObject());
  }  // end of getChartDoc()





  public static XTableChart getTableChart(XSpreadsheet sheet, 
                                               String chartName)
  // return the named table chart from the sheet
  {
    // get the supplier for the table charts
    XTableChartsSupplier chartsSupplier = 
                         Lo.qi(XTableChartsSupplier.class, sheet);
    XTableCharts tableCharts = chartsSupplier.getCharts();
    XNameAccess tcAccess = Lo.qi(XNameAccess.class, tableCharts);

    // try to access the chart with the specified name
    XTableChart tableChart = null;
    try {
      tableChart = Lo.qi(XTableChart.class, tcAccess.getByName(chartName));
      // System.out.println("Found a table chart called " + chartName);
    }
    catch(Exception ex)
    {  System.out.println("Could not access " + chartName); }
    return tableChart;
  }  // end of getTableChart()
  




  public static String[] getChartTemplates(XChartDocument chartDoc)
  // get the names of all the chart templates
  { XChartTypeManager ctMan = chartDoc.getChartTypeManager();
    return Info.getAvailableServices(ctMan);
  }



  // --------------------------- titles ------------------------



  public static void setTitle(XChartDocument chartDoc, String title)
  /* set the title, and use Arial, 14 pt style */
  {  
    // System.out.println("Chart title: \"" + title + "\"");
    XTitled xTitled = Lo.qi(XTitled.class, chartDoc);
    XTitle xtitle = createTitle(title);
    if (xtitle != null) {
      xTitled.setTitleObject(xtitle);
      setXTitleFont(xtitle, "Arial", 14);
    }
  }  // end of setTitle()


  public static XTitle createTitle(String titleString)
  {
    XTitle xtitle = Lo.createInstanceMCF(XTitle.class, "com.sun.star.chart2.Title");
    if (xtitle == null) {
      System.out.println("Unable to create xtitle interface");
      return null;
    } 

    XFormattedString xtitleStr = Lo.createInstanceMCF(XFormattedString.class, 
                                    "com.sun.star.chart2.FormattedString");
    if (xtitleStr == null) {
      System.out.println("Unable to create formatted string");
      return null;
    } 

    xtitleStr.setString(titleString);
    XFormattedString[] titleArray = new XFormattedString[] {xtitleStr};
    xtitle.setText(titleArray);
    return xtitle;
  }  // end of createTitle()



  public static void setXTitleFont(XTitle xtitle, String fontName, int ptSize)
  {
    XFormattedString[] foStrs = xtitle.getText();
    if (foStrs != null) {
      // Props.showObjProps("Title", foStrs[0]);
      Props.setProperty(foStrs[0], "CharFontName", fontName);
      Props.setProperty(foStrs[0], "CharHeight", ptSize);
    }
  }  // end of setXTitleFont()



  public static XTitle getTitle(XChartDocument chartDoc)
  { XTitled xTitled = Lo.qi(XTitled.class, chartDoc);
    return xTitled.getTitleObject();
  }


  public static void setSubtitle(XChartDocument chartDoc, String subtitle)
  /* subtitle is part of the diagram;
     set the subtitle, and use Arial, 12 pt style */
  {
    XDiagram diagram = chartDoc.getFirstDiagram();
    XTitled xTitled = Lo.qi(XTitled.class, diagram);
    XTitle xtitle = createTitle(subtitle);
    if (xtitle != null) {
      xTitled.setTitleObject(xtitle);
      setXTitleFont(xtitle, "Arial", 12);
    }
  }  // end of setSubtitle()


  public static XTitle getSubtitle(XChartDocument chartDoc)
  { 
    XDiagram diagram = chartDoc.getFirstDiagram();
    XTitled xTitled = Lo.qi(XTitled.class, diagram);
    return xTitled.getTitleObject();
  }


// ---------------- axes ----------------------


  public static XAxis getXAxis(XChartDocument chartDoc)
  {  return getAxis(chartDoc, Chart2.X_AXIS, 0);  }

  public static XAxis getYAxis(XChartDocument chartDoc)
  {  return getAxis(chartDoc, Chart2.Y_AXIS, 0);  }

  public static XAxis getXAxis2(XChartDocument chartDoc)
  {  return getAxis(chartDoc, Chart2.X_AXIS, 1);  }

  public static XAxis getYAxis2(XChartDocument chartDoc)
  {  return getAxis(chartDoc, Chart2.Y_AXIS, 1);  }


  public static XAxis getAxis(XChartDocument chartDoc, int axisVal, int idx)
  { 
    XCoordinateSystem coordSys = getCoordSystem(chartDoc);
    try {
      return coordSys.getAxisByDimension(axisVal, idx);
    }
    catch(Exception ex) {
      System.out.println("Could not get the axis");
      return null;
    }
  }  // end of getAxis()




  public static void setXAxisTitle(XChartDocument chartDoc, String title)
  {  setAxisTitle(chartDoc, title, Chart2.X_AXIS, 0);  }

  public static void setYAxisTitle(XChartDocument chartDoc, String title)
  {  setAxisTitle(chartDoc, title, Chart2.Y_AXIS, 0);  }


  public static void setXAxis2Title(XChartDocument chartDoc, String title)
  {  setAxisTitle(chartDoc, title, Chart2.X_AXIS, 1);  }

  public static void setYAxis2Title(XChartDocument chartDoc, String title)
  {  setAxisTitle(chartDoc, title, Chart2.Y_AXIS, 1);  }



  public static void setAxisTitle(XChartDocument chartDoc, String title, 
                                                   int axisVal, int idx)
  { XAxis axis = getAxis(chartDoc, axisVal, idx);
    if (axis == null)
      return;
    XTitled titledAxis = Lo.qi(XTitled.class, axis);
    XTitle xtitle = createTitle(title);
    if (xtitle != null) {
      titledAxis.setTitleObject(xtitle);
      setXTitleFont(xtitle, "Arial", 12);
    }
  }  // end of setAxisTitle()



  public static XTitle getXAxisTitle(XChartDocument chartDoc)
  {  return getAxisTitle(chartDoc, Chart2.X_AXIS, 0);  }

  public static XTitle getYAxisTitle(XChartDocument chartDoc)
  {  return getAxisTitle(chartDoc, Chart2.Y_AXIS, 0);  }


  public static XTitle getXAxis2Title(XChartDocument chartDoc)
  {  return getAxisTitle(chartDoc, Chart2.X_AXIS, 1);  }

  public static XTitle getYAxis2Title(XChartDocument chartDoc)
  {  return getAxisTitle(chartDoc, Chart2.Y_AXIS, 1);  }


  public static XTitle getAxisTitle(XChartDocument chartDoc, int axisVal, int idx)
  { XAxis axis = getAxis(chartDoc, axisVal, idx);
    if (axis == null)
      return null;
    XTitled titledAxis = Lo.qi(XTitled.class, axis);
    return titledAxis.getTitleObject();
  }  // end of getAxisTitle()



  public static void rotateXAxisTitle(XChartDocument chartDoc, int angle)
  {  rotateAxisTitle(chartDoc, Chart2.X_AXIS, 0, angle);  }

  public static void rotateYAxisTitle(XChartDocument chartDoc, int angle)
  {  rotateAxisTitle(chartDoc, Chart2.Y_AXIS, 0, angle);  }


  public static void rotateXAxis2Title(XChartDocument chartDoc, int angle)
  {  rotateAxisTitle(chartDoc, Chart2.X_AXIS, 1, angle);  }

  public static void rotateYAxis2Title(XChartDocument chartDoc, int angle)
  {  rotateAxisTitle(chartDoc, Chart2.Y_AXIS, 1, angle);  }


  public static void rotateAxisTitle(XChartDocument chartDoc, int axisVal, 
                                              int idx, int angle)
  // + angle is rotation counter-clockwise from horizontal
  { XTitle xtitle = getAxisTitle(chartDoc, axisVal, idx);
    if (xtitle != null)
      Props.setProperty(xtitle, "TextRotation", angle);
  }  // end of rotateAxisTitle()



  public static void showXAxisLabel(XChartDocument chartDoc, boolean isVisible)
  {  showAxisLabel(chartDoc, Chart2.X_AXIS, 0, isVisible);  }

  public static void showYAxisLabel(XChartDocument chartDoc, boolean isVisible)
  {  showAxisLabel(chartDoc, Chart2.Y_AXIS, 0, isVisible);  }


  public static void showXAxis2Label(XChartDocument chartDoc, boolean isVisible)
  {  showAxisLabel(chartDoc, Chart2.X_AXIS, 1, isVisible);  }

  public static void showYAxis2Label(XChartDocument chartDoc, boolean isVisible)
  {  showAxisLabel(chartDoc, Chart2.Y_AXIS, 1, isVisible);  }


  public static void showAxisLabel(XChartDocument chartDoc, int axisVal, 
                                               int idx, boolean isVisible)
  { XAxis axis = getAxis(chartDoc, axisVal, idx);
    if (axis == null)
      return;
    //Props.showObjProps("Axis", axis);
    Props.setProperty(axis, "Show", isVisible);
  }  // end of showAxisLabel()



  public static XAxis scaleXAxis(XChartDocument chartDoc, int scaleType)
  {  return scaleAxis(chartDoc, Chart2.X_AXIS, 0, scaleType);  }

  public static XAxis scaleYAxis(XChartDocument chartDoc, int scaleType)
  {  return scaleAxis(chartDoc, Chart2.Y_AXIS, 0, scaleType);  }


  public static XAxis scaleAxis(XChartDocument chartDoc, int axisVal, 
                                               int idx, int scaleType)
  /* scaleTypes: LINEAR, LOGARITHMIC, EXPONENTIAL, POWER, but
     latter two seem unstable */
  { 
    XAxis axis = getAxis(chartDoc, axisVal, idx);
    if (axis == null)
      return null;
    ScaleData sd = axis.getScaleData();
    if (scaleType == LINEAR)
      sd.Scaling = Lo.createInstanceMCF(XScaling.class,
                         "com.sun.star.chart2.LinearScaling");
    else if (scaleType == LOGARITHMIC)
      sd.Scaling = Lo.createInstanceMCF(XScaling.class,
                         "com.sun.star.chart2.LogarithmicScaling");
    else if (scaleType == EXPONENTIAL)
      sd.Scaling = Lo.createInstanceMCF(XScaling.class,
                         "com.sun.star.chart2.ExponentialScaling");
    else if (scaleType == POWER)
      sd.Scaling = Lo.createInstanceMCF(XScaling.class,
                         "com.sun.star.chart2.PowerScaling");
    else
      System.out.println("Did not recognize scaling type: " + scaleType);
    axis.setScaleData(sd);
    return axis;
  }  // end of scaleAxis()



  public static void printScaleData(String axisName, XAxis axis)
  {
    ScaleData sd = axis.getScaleData();
    System.out.println("Scaled Data for " + axisName);
    System.out.println("  Minimum: " + sd.Minimum);
    System.out.println("  Maximum: " + sd.Maximum);
    System.out.println("  Origin: " + sd.Origin);

    if (sd.Orientation == AxisOrientation.MATHEMATICAL)
      System.out.println("  Orientation: mathematical");
    else
      System.out.println("  Orientation: reverse");

    System.out.println("  Scaling: " + Info.getImplementationName(sd.Scaling));
    System.out.println("  AxisType: " + getAxisTypeString(sd.AxisType));
    System.out.println("  AutoDateAxis: " + sd.AutoDateAxis);
    System.out.println("  ShiftedCategoryPosition: " + sd.ShiftedCategoryPosition);
    System.out.println("  IncrementData: " + sd.IncrementData);
    System.out.println("  TimeIncrement: " + sd.TimeIncrement);
  }  // end of printScaleData()


  public static String getAxisTypeString(int axisType)
  {
    if (axisType == AxisType.REALNUMBER)
      return "real numbers";
    else if (axisType == AxisType.PERCENT)
      return "percentages";
    else if (axisType == AxisType.CATEGORY)
      return "categories";
    else if (axisType == AxisType.SERIES)
      return "series names";
    else if (axisType == AxisType.DATE)
      return "dates";
    else
      return "unknown";
  }  // end of getAxisTypeString()



// ---------------- grid lines ----------------------


  public static void setGridLines(XChartDocument chartDoc, int axisVal)
  {  setGridLines(chartDoc, axisVal, 0);  }


  public static void setGridLines(XChartDocument chartDoc, int axisVal, int idx)
  { 
    XAxis axis = getAxis(chartDoc, axisVal, idx);   // only go for major 
    if (axis == null)
      return;
    XPropertySet props = axis.getGridProperties();
    // Props.showProps("Axis Major Grid", props);
    Props.setProperty(props, "LineStyle", LineStyle.DASH);
    Props.setProperty(props, "LineDashName", LINE_STYLES[3]);  // "Fine Dotted"
  }  // end of setGridLines()



  // ------------------------- legend ----------------------


  public static void viewLegend(XChartDocument chartDoc, boolean isVisible)
  {
    XDiagram diagram = chartDoc.getFirstDiagram();
    XLegend legend = diagram.getLegend();
    if (isVisible && (legend == null)) {
      XLegend leg = Lo.createInstanceMCF(XLegend.class, "com.sun.star.chart2.Legend");
      Props.setProperty(leg, "LineStyle", LineStyle.NONE);
      Props.setProperty(leg, "FillStyle", FillStyle.SOLID);
      Props.setProperty(leg, "FillTransparence", 100);  // transparent
      diagram.setLegend(leg);
    }
    Props.setProperty(diagram.getLegend(), "Show", isVisible);  // toggle visibility
  }  // end of viewLegend()



  // ------------------------- background colors ----------------------


  public static void setBackgroundColors(XChartDocument chartDoc, 
                                              int bgColor, int wallColor)
  {
    if (bgColor > 0) {
      XPropertySet bgProps = chartDoc.getPageBackground();
      // Props.showProps("Background", bgProps);
      Props.setProperty(bgProps, "FillBackground", true);
      Props.setProperty(bgProps, "FillStyle", FillStyle.SOLID);
      Props.setProperty(bgProps, "FillColor", bgColor);
    }

    if (wallColor > 0) {
      XDiagram diagram = chartDoc.getFirstDiagram();
      XPropertySet wallProps = diagram.getWall();
      // Props.showProps("Wall", wallProps);
      Props.setProperty(wallProps, "FillBackground", true);
      Props.setProperty(wallProps, "FillStyle", FillStyle.SOLID);
      Props.setProperty(wallProps, "FillColor", wallColor);
    }
  }  // end of setBackgroundColors()




  // -------------- access data source and series -------------------------



  public static XDataSource getDataSource(XChartDocument chartDoc)
  {
    XDataSeries[] dsa = getDataSeries(chartDoc); 
    // System.out.println("No. of data series: " + dsa.length);
    // Info.showInterfaces("Data Series Array", dsa);
    return Lo.qi(XDataSource.class, dsa[0]);
  }  // end of getDataSource()


  public static XDataSource getDataSource(XChartDocument chartDoc, String chartType)
  {
    XDataSeries[] dsa = getDataSeries(chartDoc, chartType); 
    return Lo.qi(XDataSource.class, dsa[0]);
  }  // end of getDataSource()




  public static XDataSeries[] getDataSeries(XChartDocument chartDoc)
  // get the data series associated with the first chart type
  {
    XChartType xChartType = getChartType(chartDoc);
    XDataSeriesContainer dsCon = 
                     Lo.qi(XDataSeriesContainer.class, xChartType);
    return dsCon.getDataSeries();
  }  //end of getDataSeries()


  public static XDataSeries[] getDataSeries(XChartDocument chartDoc, String chartType)
  // get the data series associated with the specified chart type
  {
    XChartType xChartType = findChartType(chartDoc, chartType);
    if (xChartType == null)
      return null;
    XDataSeriesContainer dsCon = 
                     Lo.qi(XDataSeriesContainer.class, xChartType);
    return dsCon.getDataSeries();
  }  //end of getDataSeries()




  public XDataSeries createDataSeries()
  {
    XDataSeries ds = Lo.createInstanceMCF(XDataSeries.class, "com.sun.star.chart2.DataSeries");
    if (ds == null)
      System.out.println("Unable to create XDataSeries interface: " + ds);
    return ds;
  }




  // ------------------------ chart types -------------------------



  public static XChartType getChartType(XChartDocument chartDoc)
  {
    XChartType[] chartTypes = getChartTypes(chartDoc);
    return chartTypes[0];
  }


  public static XChartType[] getChartTypes(XChartDocument chartDoc)
  {
    XCoordinateSystem coordSys = getCoordSystem(chartDoc);
    XChartTypeContainer ctCon = Lo.qi(XChartTypeContainer.class, coordSys);
    return ctCon.getChartTypes();
  }  // end of getChartTypes()



  public static XCoordinateSystem getCoordSystem(XChartDocument chartDoc)
  {
    XDiagram diagram = chartDoc.getFirstDiagram();
    XCoordinateSystemContainer coordSysCon = 
                     Lo.qi(XCoordinateSystemContainer.class, diagram);
    XCoordinateSystem[] coordSys = coordSysCon.getCoordinateSystems();
    if (coordSys.length > 1)
      System.out.println("No of coord systems: " + coordSys.length + "; using first");
    return coordSys[0];
  }  // end of getCoordSystem()




  public static void printChartTypes(XChartDocument chartDoc)
  {
    XChartType[] chartTypes = getChartTypes(chartDoc);
    if (chartTypes.length > 1) {
      System.out.println("No. of chart types: " + chartTypes.length);
      for(XChartType ct : chartTypes)
        System.out.println("  " + ct.getChartType());
    }
    else
      System.out.println("Chart type: " + chartTypes[0].getChartType());
  }  // end of printChartTypes()



  public static XChartType findChartType(XChartDocument chartDoc, String chartType)
  {
    String srchName = "com.sun.star.chart2." + chartType.toLowerCase();
    XChartType[] chartTypes = getChartTypes(chartDoc);
    for(XChartType ct : chartTypes) {
      String ctName = ct.getChartType().toLowerCase();
      if (ctName.equals(srchName))
        return ct;
    }
    System.out.println("Chart type " + srchName + " not found");
    return null;
  }  // end of findChartType()



  public static XChartType addChartType(XChartDocument chartDoc, String chartType)
  {
    XChartType ct = Lo.createInstanceMCF(XChartType.class, 
                                             "com.sun.star.chart2." + chartType);
    if (ct == null) {
      System.out.println("Unable to create XChartType interface: " + chartType);
      return ct;
    }
 
    XCoordinateSystem coordSys = getCoordSystem(chartDoc);
    XChartTypeContainer ctCon = Lo.qi(XChartTypeContainer.class, coordSys);
    ctCon.addChartType(ct);
    return ct;
  }  // end of addChartType()



// ----------------------- using a data source -----------------------



  public static void showDataSourceArgs(XChartDocument chartDoc, XDataSource dataSource)
  {
    XDataProvider dp = chartDoc.getDataProvider();
    //Props.showObjProps("Data Source", dataSource);
    PropertyValue[] props = dp.detectArguments(dataSource); 
    Props.showProps("Data Source arguments", props);
  }  // end of showDSArgs()



  public static void printLabeledSeqs(XDataSource dataSource)
  { 
    XLabeledDataSequence[] dataSeqs = dataSource.getDataSequences();
    System.out.println("No. of sequences in data source: " + dataSeqs.length);
    for (int i=0; i < dataSeqs.length; i++) {
      Object[] labelSeq = dataSeqs[i].getLabel().getData();
      System.out.print(labelSeq[0] + " :");
      Object[] valsSeq = dataSeqs[i].getValues().getData();
      for (Object val : valsSeq)
        System.out.print("  " + val);
      System.out.println();
      String srRep = dataSeqs[i].getValues().getSourceRangeRepresentation();
      System.out.println("  Source range: " + srRep);
    }
  }  // end of printLabeledSeqs()



  public static double[] getChartData(XDataSource dataSource, int seqIdx)
  {  
    XLabeledDataSequence[] dataSeqs = dataSource.getDataSequences();
    if ((seqIdx < 0) || (seqIdx >= dataSeqs.length)) {
      System.out.println("Index is out of range");
      return null;
    }

    Object[] valsSeq = dataSeqs[seqIdx].getValues().getData();
    double[] vals = new double[valsSeq.length];
    for (int i=0; i < valsSeq.length; i++)
       vals[i] = (double) valsSeq[i];
    return vals;
  }  // end of getChartData()




  // --------- using data series point props ------------



  public static XPropertySet getDataPointProps(XChartDocument chartDoc,
                                                      int seriesIdx, int idx)
  { 
    XPropertySet[] propsArr = getDataPointsProps(chartDoc, seriesIdx);
    if (propsArr == null)
      return null;

    if ((idx < 0) || (idx >= propsArr.length)) {
      System.out.println("No data at index " + idx + "; use 0 to " + (propsArr.length-1));
      return null;
    }

    return propsArr[idx]; 
  }  // end getDataPointProps()



  public static XPropertySet[] getDataPointsProps(XChartDocument chartDoc, int seriesIdx)
  // get all the properties for the data in the specified series
  {
    XDataSeries[] dataSeriesArr = getDataSeries(chartDoc);
    // System.out.println("Size of data series array: " + dataSeriesArr.length);

    if ((seriesIdx < 0) || (seriesIdx >= dataSeriesArr.length)) {
      System.out.println("Series index is out of range");
      return null;
    }
 
    ArrayList<XPropertySet> propsList = new ArrayList<XPropertySet>();
    int idx = 0;
    XPropertySet props = null;
    do {
      try {
        props = dataSeriesArr[seriesIdx].getDataPointByIndex(idx++);
        if (props != null)
          propsList.add(props); 
      }
      catch(com.sun.star.lang.IndexOutOfBoundsException e) { 
        break;  
      }
    } while (props != null);

    if (propsList.size() == 0) {
      System.out.println("No Series at index " + seriesIdx);
      return null;
    }
    XPropertySet[] propsArr = new XPropertySet[propsList.size()];
    for (int i=0; i < propsList.size(); i++)
       propsArr[i] = propsList.get(i);
    return propsArr;
  }  // end getDataPointsProps()





  public static void setDataPointLabels(XChartDocument chartDoc, int labelType)
  // labeltype can be DP_NUMBER, DP_PERCENT, DP_CATEGORY, DP_SYMBOL, DP_NONE
  {
    XDataSeries[] dataSeriesArr = getDataSeries(chartDoc);
    // System.out.println("No. of data series: " + dataSeriesArr.length);
    // Props.showObjProps("Data Series 0", dataSeriesArr[0]);

    for (XDataSeries dataSeries : dataSeriesArr) {
      DataPointLabel dpLabel = 
               (DataPointLabel) Props.getProperty(dataSeries, "Label");
      dpLabel.ShowNumber = false;
      dpLabel.ShowCategoryName = false;
      dpLabel.ShowLegendSymbol = false;
      if (labelType == DP_NUMBER)
        dpLabel.ShowNumber = true;
      else if (labelType == DP_PERCENT) {
        dpLabel.ShowNumber = true;
        dpLabel.ShowNumberInPercent = true;
      }
      else if (labelType == DP_CATEGORY)
        dpLabel.ShowCategoryName = true;
      else if (labelType == DP_SYMBOL)
        dpLabel.ShowLegendSymbol = true;
      else if (labelType == DP_NONE) {}
      else
        System.out.println("Unrecognized label type");
      Props.setProperty(dataSeries, "Label", dpLabel);
    }
  }  // end of setDataPointLabels()




  public static void setChartShape3D(XChartDocument chartDoc, String shape)
  { 
    XDataSeries[] dataSeriesArr = getDataSeries(chartDoc);
    for (XDataSeries dataSeries : dataSeriesArr) {
      if (shape.equals("box"))
        Props.setProperty(dataSeries, "Geometry3D", 
                                      DataPointGeometry3D.CUBOID);
      else if (shape.equals("cylinder"))
        Props.setProperty(dataSeries, "Geometry3D", 
                                      DataPointGeometry3D.CYLINDER);
      else if (shape.equals("cone"))
        Props.setProperty(dataSeries, "Geometry3D", 
                                      DataPointGeometry3D.CONE);
      else if (shape.equals("pyramid"))
        Props.setProperty(dataSeries, "Geometry3D", 
                                      DataPointGeometry3D.PYRAMID);
      else
        System.out.println("Did not recognise 3D shape: " + shape);
    }
  }  // end of setChartShape3D()




  public static void dashLines(XChartDocument chartDoc)
  {
    XDataSeries[] dataSeriesArr = getDataSeries(chartDoc);
    for (XDataSeries dataSeries : dataSeriesArr) {
      // Props.setProperty(dataSeries, "LineDash", dash);
      Props.setProperty(dataSeries, "LineStyle", LineStyle.DASH);
      Props.setProperty(dataSeries, "LineDashName", LINE_STYLES[1]);  // "Fine Dashed"
             // also LineWidth, LineStyle
    }
  }  // end of dashLines()



  public static void colorStockBars(XChartType ct, int wDayColor, int bDayColor)
  /* Change the WhiteDay and BlackDay colors in the stock bars.
     WhiteDay is usually white and means an increase; BlackDay is usually
     black and means a decrease.
  */
  {
    if (!ct.getChartType().equals("com.sun.star.chart2.CandleStickChartType"))
      System.out.println("Chart type not a candle stick: " + ct.getChartType());
    else {
      XPropertySet props = Lo.qi(XPropertySet.class, Props.getProperty(ct, "WhiteDay"));
      // Props.showObjProps("WhiteDay", props);
      Props.setProperty(props, "FillColor", wDayColor);

      props = Lo.qi(XPropertySet.class, Props.getProperty(ct, "BlackDay"));
      Props.setProperty(props, "FillColor", bDayColor);
    }
  }  // end of colorStockBars()



  // ----------------------- regression -------------------------



  public static void drawRegressionCurve(XChartDocument chartDoc, int curveKind)
  {
    XDataSeries[] dataSeriesArr = getDataSeries(chartDoc);
    // System.out.println("No of data series: " + dataSeriesArr.length);

    XRegressionCurveContainer rcCon = Lo.qi(
                         XRegressionCurveContainer.class, dataSeriesArr[0]);
    // XRegressionCurve[] curves = rcCon.getRegressionCurves();

    XRegressionCurve curve = createCurve(curveKind);
    rcCon.addRegressionCurve(curve);
    // Props.showObjProps("Regression curve", curve);

    // show equation and R^2 value
    XPropertySet props = curve.getEquationProperties();
    Props.setProperty(props, "ShowCorrelationCoefficient", true);
    Props.setProperty(props, "ShowEquation", true);

    int key = getNumberFormatKey(chartDoc, "0.00");   // 2 dp
    if (key != -1)
      Props.setProperty(props, "NumberFormat", key);

    // Props.showProps("Regression Equation", props);
  }  // end of drawRegressionCurve()



  public static XRegressionCurve createCurve(int curveKind)
  {
    if (curveKind == LINEAR)
      return Lo.createInstanceMCF(XRegressionCurve.class, 
                  "com.sun.star.chart2.LinearRegressionCurve");
    else if (curveKind == LOGARITHMIC)
      return Lo.createInstanceMCF(XRegressionCurve.class, 
                  "com.sun.star.chart2.LogarithmicRegressionCurve");
    else if (curveKind == EXPONENTIAL)
      return Lo.createInstanceMCF(XRegressionCurve.class, 
                  "com.sun.star.chart2.ExponentialRegressionCurve");
    else if (curveKind == POWER)
      return Lo.createInstanceMCF(XRegressionCurve.class, 
                  "com.sun.star.chart2.PotentialRegressionCurve");
    else if (curveKind == POLYNOMIAL)         // assume degree == 2
      return Lo.createInstanceMCF(XRegressionCurve.class, 
                   "com.sun.star.chart2.PolynomialRegressionCurve");
    else if (curveKind == MOVING_AVERAGE)    // assume period == 2
      return Lo.createInstanceMCF(XRegressionCurve.class, 
                   "com.sun.star.chart2.MovingAverageRegressionCurve");
    else {   
      System.out.println("Did not recognize regression line kind: " + curveKind + "; using linear");
      return Lo.createInstanceMCF(XRegressionCurve.class, 
                  "com.sun.star.chart2.LinearRegressionCurve");
    }
  }  // end of createCurve()



  public static int getNumberFormatKey(XChartDocument chartDoc, String nfStr)
  {
    XNumberFormatsSupplier xfs =  
                  Lo.qi(XNumberFormatsSupplier.class, chartDoc);
    XNumberFormats nFormats = xfs.getNumberFormats();
    int key = (int) nFormats.queryKey(nfStr, 
                new com.sun.star.lang.Locale("en", "us", ""), false);
    if (key == -1)
      System.out.println("Could not access key for number format: \"" + 
                                                           nfStr + "\"");
    return key;
  }  // end of getNumberFormatKey()



  public static void calcRegressions(XChartDocument chartDoc)
  {
    for(int i=0; i < CURVE_KINDS.length; i++) {
      XRegressionCurve curve = createCurve(CURVE_KINDS[i]);
      System.out.println(CURVE_NAMES[i] + " regression curve:");
      evalCurve(chartDoc, curve);
      System.out.println();
    }
  }  // end of calcRegressions()


  public static void evalCurve(XChartDocument chartDoc, XRegressionCurve curve)
  {
    XRegressionCurveCalculator curveCalc = curve.getCalculator();

    int degree = 1;
    if (getCurveType(curve) != LINEAR)
      degree = 2;   // assumes POLYNOMIAL trend has degree == 2
    curveCalc.setRegressionProperties(degree, false, 0, 2);
      //  degree, forceIntercept, interceptValue, period (for moving average)


    XDataSource dataSource = getDataSource(chartDoc);
    // printLabeledSeqs(dataSource);

    double[] xVals = getChartData(dataSource, 0);   // assuming xy scatter chart
    double[] yVals = getChartData(dataSource, 1);
    curveCalc.recalculateRegression(xVals, yVals);

    System.out.println("  Curve equation: " + curveCalc.getRepresentation());
    double cc = curveCalc.getCorrelationCoefficient();
    System.out.printf("  R^2 value: %.3f\n", (cc*cc));   // 3 dp
  }  // end of evalCurve()



  public static int getCurveType(XRegressionCurve curve)
  {
    String[] services = Info.getServices(curve);
    if (Arrays.binarySearch(services, 
               "com.sun.star.chart2.LinearRegressionCurve") >= 0)
      return LINEAR;
    else if (Arrays.binarySearch(services, 
               "com.sun.star.chart2.LogarithmicRegressionCurve") >= 0)
      return LOGARITHMIC;
    else if (Arrays.binarySearch(services, 
               "com.sun.star.chart2.ExponentialRegressionCurve") >= 0)
      return EXPONENTIAL;
    else if (Arrays.binarySearch(services, 
               "com.sun.star.chart2.PotentialRegressionCurve") >= 0)
      return POWER;
    else if (Arrays.binarySearch(services, 
               "com.sun.star.chart2.PolynomialRegressionCurve") >= 0)
      return POLYNOMIAL;
    else if (Arrays.binarySearch(services, 
               "com.sun.star.chart2.MovingAverageRegressionCurve") >= 0)
      return MOVING_AVERAGE;
    else {
      System.out.println("Could not identify trend type of curve; using linear");
      return LINEAR;
    }
  }  // end of getCurveType()




  // -------------------- add data to a chart ---------------------------


  public static void setYErrorBars(XChartDocument chartDoc, 
                                  String dataLabel, String dataRange)
  {
    // initialize error bar properties
    XPropertySet errorBarProps = 
        Lo.createInstanceMCF(XPropertySet.class, 
                             "com.sun.star.chart2.ErrorBar");
    Props.setProperty(errorBarProps, "ShowPositiveError", true);
    Props.setProperty(errorBarProps, "ShowNegativeError", true);
    Props.setProperty(errorBarProps, "ErrorBarStyle", 
                                         ErrorBarStyle.FROM_DATA);

    // convert into data sink
    XDataSink dataSink = Lo.qi(XDataSink.class, errorBarProps);


    // use data provider to create labelled data sequences
    // for the +/- error ranges
    // System.out.println("Error range: " + dataRange);
    XDataProvider dp = chartDoc.getDataProvider();

    XLabeledDataSequence posErrSeq = 
            createLDSeq(dp, "error-bars-y-positive", dataLabel, dataRange);
    XLabeledDataSequence negErrSeq = 
            createLDSeq(dp, "error-bars-y-negative", dataLabel, dataRange);
    XLabeledDataSequence[] ldSeqArr = { posErrSeq, negErrSeq };

    // store the error bar data sequences in the data sink
    dataSink.setData(ldSeqArr);
    // Props.showObjProps("Error Bar", errorBarProps);
       // "ErrorBarRangePositive" and "ErrorBarRangeNegative" 
       // will now have ranges they are read-only

    // store error bar in data series
    XDataSeries[] dataSeriesArr = getDataSeries(chartDoc);
    // System.out.println("No. of data series: " + dataSeriesArr.length);
    XDataSeries dataSeries = dataSeriesArr[0];
    // Props.showObjProps("Data Series 0", dataSeries);
    Props.setProperty(dataSeries, "ErrorBarY", errorBarProps);
  }  // end of setYErrorBars()



  public static XLabeledDataSequence createLDSeq(XDataProvider dp, 
                      String role, String dataLabel, String dataRange)
  // create labeled data sequence using label and data;
  // the data is for the specified role
  {
    // create data sequence for the label
    XDataSequence labelSeq =  
                dp.createDataSequenceByRangeRepresentation(dataLabel);

    // create data sequence for the data and role
    XDataSequence dataSeq =  
              dp.createDataSequenceByRangeRepresentation(dataRange);
    XPropertySet dsProps = Lo.qi(XPropertySet.class, dataSeq);
    Props.setProperty(dsProps, "Role", role);  //specify data role (type)
    // Props.showObjProps("Data Sequence", dsProps);

    // create new labeled data sequence using sequences
    XLabeledDataSequence ldSeq =  
                Lo.createInstanceMCF(XLabeledDataSequence.class, 
                      "com.sun.star.chart2.data.LabeledDataSequence");
    ldSeq.setLabel(labelSeq);  // add label
    ldSeq.setValues(dataSeq);  // add data
    return ldSeq;
  }  // end of createLDSeq()



  public static void addStockLine(XChartDocument chartDoc, 
                                  String dataLabel, String dataRange)
  {
    // add (empty) line chart to the doc
    XChartType ct = addChartType(chartDoc, "LineChartType");
    XDataSeriesContainer dataSeriesCnt = Lo.qi(XDataSeriesContainer.class, ct);

    // create (empty) data series in the line chart
    XDataSeries ds = Lo.createInstanceMCF(XDataSeries.class, "com.sun.star.chart2.DataSeries");
    if (ds == null) {
      System.out.println("Unable to create XDataSeries interface: " + ds);
      return;
    }
    Props.setProperty(ds, "Color", 0xFF0000);
    dataSeriesCnt.addDataSeries(ds);

    
    // add data to series by treating it as a data sink
    XDataSink dataSink = Lo.qi(XDataSink.class, ds);

    // add data as y values
    XDataProvider dp = chartDoc.getDataProvider();
    XLabeledDataSequence dLSeq = createLDSeq(dp, "values-y", dataLabel, dataRange);
    XLabeledDataSequence[] ldSeqArr = { dLSeq };
    dataSink.setData(ldSeqArr);

  }  // end of addStockLine()



  public static void addCatLabels(XChartDocument chartDoc, 
                                  String dataLabel, String dataRange)
  {
    // add data as categories for the x-axis
    XDataProvider dp = chartDoc.getDataProvider();
    XLabeledDataSequence dLSeq = createLDSeq(dp, "categories", dataLabel, dataRange);

    XAxis axis = getAxis(chartDoc, X_AXIS, 0);
    if (axis == null)
      return;
    ScaleData sd = axis.getScaleData();
    sd.Categories = dLSeq;
    axis.setScaleData(sd);

    setDataPointLabels(chartDoc, Chart2.DP_CATEGORY);
       // label the data points with these category values
  }  // end of addCatLabels()




  // --------------------- chart shape and image ------------------



  public static void copyChart(XSpreadsheetDocument ssdoc, 
                                         XSpreadsheet sheet)
  {
    XShape chartShape = getChartShape(sheet);
    XComponent doc = Lo.qi(XComponent.class, ssdoc);
    XSelectionSupplier supplier = GUI.getSelectionSupplier(doc);
    supplier.select((Object)chartShape);
    Lo.dispatchCmd("Copy"); 
  }  // end of copyChart()




  public static XShape getChartShape(XSpreadsheet sheet)
  // return the first chart shape
  {
    //get draw page supplier for chart sheet 
    XDrawPageSupplier pageSupplier = Lo.qi(XDrawPageSupplier.class, sheet); 
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
        if (classID.toLowerCase().equals(Lo.CHART_CLSID))
          break;
      }
      catch(Exception e) {}
    }
    //if (shape != null)
    //  System.out.println("Found a chart");
    return shape;
  }  // end of getChartShape()



  public static BufferedImage getChartImage(XSpreadsheet sheet)
  {
    XShape chartShape = Chart2.getChartShape(sheet);
    if (chartShape == null) {
      System.out.println("Could not find a chart");
      return null;
    }

    //System.out.println("Shape type: " + chartShape.getShapeType());
    //Props.showObjProps("Shape", chartShape);
    XGraphic graphic = Lo.qi( XGraphic.class, 
                             Props.getProperty(chartShape, "Graphic") );
    if (graphic == null) {
      System.out.println("No chart graphic found");
      return null;
    }

    String tempFnm = FileIO.createTempFile("png");
    if (tempFnm == null) {
      System.out.println("Could not create a temporary file for the graphic");
      return null;
    }

    Images.saveGraphic(graphic, tempFnm, "png");
    BufferedImage im = Images.loadImage(tempFnm);
    FileIO.deleteFile(tempFnm);
    return im;
  }  // end of getChartImage()




  public static XDrawPage getChartDrawPage(XSpreadsheet sheet)
  {
    XShape chartShape = getChartShape(sheet);

    XEmbeddedObject embeddedChart =  Lo.qi(XEmbeddedObject.class,  
                      Props.getProperty(chartShape, "EmbeddedObject"));
    XComponentSupplier xComponentSupplier = 
               Lo.qi(XComponentSupplier.class, embeddedChart);
    XCloseable xCloseable = xComponentSupplier.getComponent();
    XDrawPageSupplier xSuppPage = Lo.qi(XDrawPageSupplier.class, xCloseable); 
    return xSuppPage.getDrawPage(); 
  }  // end of getChartDrawPage()





}  // end of Chart2 class


// Base.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, February 2016

/* A growing collection of utility functions to make Office Base
   easier to use. They are currently divided into the following
   groups:

     * create, open, close databases
     * execute methods: query, update

     * display databases & tables
     * display result sets
     * save a database as CSV files

     * read in SQL commands
     * zip-related methods for an ODB file
     * extract embedded DB files
     * rowset method

     * database/table info using SDBCX
     * database/table info using SDBC
     * data source info
     * DriverManager and Drives info

  Utilities for using JDBC are in Jdbc.java
*/

package utils;

import com.sun.star.beans.*;
import com.sun.star.container.*;
import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;
import com.sun.star.util.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.XCloseable;
import com.sun.star.sdbc.*;
import com.sun.star.sdbcx.*;

import com.sun.star.sdb.application.*;
import com.sun.star.form.runtime.*;


import java.io.*;
import java.awt.*;
import java.util.*;
import java.util.zip.*;
import javax.swing.*;
import javax.swing.table.*;


import javax.xml.parsers.*;
import javax.xml.xpath.*;
import org.w3c.dom.*;
import org.xml.sax.SAXException;




public class Base
{
  public static final String TEMP_DIR = "baseTmp/";

  // used in createBaseDoc() for *embedded* databases of these types
  public static final int UNKNOWN = 0;
  public static final int HSQLDB = 1;
  public static final int FIREBIRD = 2;

  public static final String HSQL_EMBEDDED = "sdbc:embedded:hsqldb";
  public static final String FIREBIRD_EMBEDDED = "sdbc:embedded:firebird";

  // names used in zipped ODB file structure
  public static final String HSQL_FNM = "hsqlDatabase";
  public static final String FB_FNM = "firebird";
  public static final String ZIP_DIR_NM = "database/";



  // -------------  create, open, close databases -------------------


  public static XOfficeDatabaseDocument createBaseDoc(
                              String fnm, XComponentLoader loader)
  {  return createBaseDoc(fnm, Base.HSQLDB, loader);  }



  public static XOfficeDatabaseDocument createBaseDoc(
                         String fnm, int dbType, XComponentLoader loader)
  {  
    if ((dbType != HSQLDB) && (dbType != FIREBIRD)) {
      System.out.println("Did not recognize the database type; using HSQLDB");
      dbType = HSQLDB;
    }

    XComponent doc = Lo.createDoc("sdatabase", loader);
    if (doc == null) {
      System.out.println("Document creation failed");
      return null;
    }

    XOfficeDatabaseDocument dbDoc = Lo.qi(XOfficeDatabaseDocument.class, doc);
    XDataSource dataSource = dbDoc.getDataSource();

    String sourceStr = (dbType == FIREBIRD) ? FIREBIRD_EMBEDDED : HSQL_EMBEDDED;
    Props.setProperty(dataSource, "URL", sourceStr);

    System.out.println("Created a \"" + sourceStr + "\" Base document in " + fnm);
    
    // must save, close, reopen, or no XConnection can be made to database
    Base.saveBaseDoc(dbDoc, fnm);
    Base.closeBaseDoc(dbDoc);
    dbDoc = openBaseDoc(fnm, loader);
    System.out.println();
    // System.out.println("Database type: " + Base.getDataSourceType(dbDoc));
    
    return dbDoc;
  }  // end of createBaseDoc()




  public static void saveBaseDoc(XOfficeDatabaseDocument dbDoc, String fnm)
  { // XStorable store = Lo.qi(XStorable.class, dbDoc);
    XComponent doc = Lo.qi(XComponent.class, dbDoc);
    Lo.saveDoc(doc, fnm);
  }



  public static void closeConnection(XConnection conn)
  {
    if (conn == null)
      return;
    try {
      XCloseable closeConn = Lo.qi(XCloseable.class, conn);
      if (closeConn != null) {
        try {
          closeConn.close();
          if (closeConn != null) {
            XComponent xc = Lo.qi(XComponent.class, conn);
            if (xc != null)
              xc.dispose();
          }
          System.out.println("Closed database connection");
        }
        catch (SQLException e) 
        {  System.out.println("Unable to close database connection");  }
      }
    }
    catch (com.sun.star.lang.DisposedException e) {
       System.out.println("Database connection close failed since Office link disposed");
    }
  }  // end of closeConnection()



  public static void closeBaseDoc(XOfficeDatabaseDocument dbDoc)
  { 
    try {
      com.sun.star.util.XCloseable closeable = Lo.qi(
                   com.sun.star.util.XCloseable.class, dbDoc);
      Lo.close(closeable);
    }
    catch (com.sun.star.lang.DisposedException e) {
       System.out.println("Database close failed since Office link disposed");
    }
  }  // end of closeBaseDoc()




  public static XOfficeDatabaseDocument openBaseDoc(String fnm, XComponentLoader loader)
  {
    XComponent doc = Lo.openDoc(fnm, loader);
    if (doc == null) {
      System.out.println("Document is null");
      return null;
    }

    if (Info.reportDocType(doc) != Lo.BASE) {
      System.out.println("Not a database doc; closing " + fnm);
      Lo.closeDoc(doc);
      return null;
    }

    XOfficeDatabaseDocument dbDoc = Lo.qi(XOfficeDatabaseDocument.class, doc);
    if (dbDoc == null) {
      System.out.println("Not a database doc; closing " + fnm);
      Lo.closeDoc(doc);
      return null;
    }

    return dbDoc;
  }  // end of openDoc()



  public static void refreshTables(XConnection conn)
  {
    XTablesSupplier suppTables = Lo.qi(XTablesSupplier.class, conn);
    XRefreshable refresh = Lo.qi(XRefreshable.class, suppTables.getTables());
    refresh.refresh();
  }  // end of refreshTables()




  public static void showTables(XOfficeDatabaseDocument dbDoc)
  /* This way of showing tables often causes Office to crash
     when Office is closed while the table windows are open
     on-screen.

     It's probably better to use Base.printDatabase() or
     Base.displayDatabase()
  */
  {
    showTablesView(dbDoc);
    Lo.delay(500);  // wait for Tables View to appear
    Lo.dispatchCmd("SelectAll");
    Lo.dispatchCmd("DBTableOpen");  // open all tables
  }  // end of showTables()


  public static void showTablesView(XOfficeDatabaseDocument dbDoc)
  {
    XComponent doc = Lo.qi(XComponent.class, dbDoc);
    GUI.setVisible(doc, true);
    Lo.delay(500);  // wait for GUI to appear
    Lo.dispatchCmd("DBViewTables");
  }  // end of showTablesView()



  // --------------------- use DB document UI -------------------

  public static XDatabaseDocumentUI getDUI(XOfficeDatabaseDocument dbDoc)
  {
    XController ctrl = GUI.getCurrentController(dbDoc);
    return Lo.qi(XDatabaseDocumentUI.class, ctrl);
  }



  public static XController loadTable(XOfficeDatabaseDocument dbDoc, String tableName)
  {
    try {
      XDatabaseDocumentUI docUI = getDUI(dbDoc);
      if (!docUI.isConnected()) {
        System.out.println("Database not connected");
        docUI.connect();
      }
      XComponent tableComp = docUI.loadComponent(DatabaseObject.TABLE, tableName, false);
      XController ctrl = Lo.qi(XController.class, tableComp);
      if (ctrl != null)
         return ctrl;
      XModel model = Lo.qi(XModel.class, tableComp);
      return model.getCurrentController();
    }
    catch(com.sun.star.uno.Exception e)
    {  e.printStackTrace();
       // System.out.println(e);  
       return null;
    }
  }  // end of loadTable()



  public static void showTable(XOfficeDatabaseDocument dbDoc, String tableName)
  {
    XController ctrl = loadTable(dbDoc, tableName);
    if (ctrl != null) {
      XFormController tableViewController = Lo.qi(XFormController.class, ctrl);

      XPropertySet props = Lo.qi(XPropertySet.class,
            tableViewController.getCurrentControl().getModel());
    }
  }  // end of showTable()




  // ------- execute methods: query, update ----------------------


  public static boolean exec(String cmd, XConnection conn)
  {
    if (cmd.startsWith("CREATE"))
      return execute(cmd, conn);
    else if (cmd.startsWith("SELECT")) {
      XResultSet rs = executeQuery(cmd, conn);
      printResultSet(rs);
      return (rs == null);
    }
    else if (cmd.startsWith("INSERT") || cmd.startsWith("UPDATE") ||
             cmd.startsWith("DELETE")) {
      int res = executeUpdate(cmd, conn);
      return (res != -1);
    }
    else {
      System.out.println("Cannot process cmd: \"" + cmd + "\"");
      return false;
    }
  }  // end of exec()



  public static boolean execute(String stmtStr, XConnection conn)
  {
    if (conn == null) {
      System.out.println("Connection is null");
      return false;
    }

    try { 
      XStatement statement = conn.createStatement();
      statement.execute(stmtStr);
      System.out.println("Executed \"" + stmtStr + "\"");
      return true;
    }
    catch(SQLException e) {
      System.out.println("Unable to execute: \"" + stmtStr + "\"\n   " + e.getMessage());
      return false;
    }
  }  // end of execute()



  public static XResultSet executeQuery(String query, XConnection conn)
  {
    if (conn == null) {
      System.out.println("Connection is null");
      return null;
    }

    try { 
      XStatement statement = conn.createStatement();
      XPropertySet xProp = Lo.qi(XPropertySet.class, statement);
      XResultSet rs = statement.executeQuery(query);
      System.out.println("Processed \"" + query + "\": " + (rs != null));
      return rs;
    }
    catch(com.sun.star.uno.Exception e) {
      System.out.println("Unable to execute query: \"" + query + "\"\n   " + e.getMessage());
      return null;
    }
  }  // end of executeQuery()



  public static int executeUpdate(String query, XConnection conn)
  {
    if (conn == null) {
      System.out.println("Connection is null");
      return -1;
    }

    try { 
      XStatement statement = conn.createStatement();
      int res = statement.executeUpdate(query);
      System.out.println("Result for \"" + query + "\": " + res);
      return res;
    }
    catch(SQLException e) {
      System.out.println("Unable to execute update: \"" + 
                                     query + "\"\n   " + e.getMessage());
      return -1;
    }
  }  // end of executeUpdate()


  // ------------- display databases & tables ---------------------
  // as printed text (using BaseTablePrinter class), as a JTable



  public static void printDatabase(XConnection conn)
  {
    if (conn == null) {
      System.out.println("Connection is null");
      return;
    }

    ArrayList<String> tableNames = getTablesNames(conn);
    if (tableNames == null)
      System.out.println("No tables found in database");
    else {
      for(String tableName : tableNames)
        BaseTablePrinter.printTable(conn, tableName);
    }
  }  // end of saveDatabase()




  public static void displayDatabase(XConnection conn)
  {
    if (conn == null) {
      System.out.println("Connection is null");
      return;
    }

    ArrayList<String> tableNames = getTablesNames(conn);
    if (tableNames == null)
      System.out.println("No tables found in database");
    else {
      int offset = 0;
      for(String tableName : tableNames) {
        displayTable(tableName, conn, offset);
        offset += 30;
      }
    }
  }  // end of displayDatabase()



  public static void displayTable(String tableName, XConnection conn, int offset)
  {
    try { 
      XStatement statement = conn.createStatement();
      XResultSet rs = statement.executeQuery("SELECT * FROM " + tableName);
      displayResultSet(rs, tableName, offset);
    }
    catch(SQLException e) {
      System.out.println(tableName + " display error: " + e);
    }
  }  // end of displayTable()



  // ------------- display result sets ---------------------
  // as printed text, as a JTable, as a 2D array


  public static int getColumnCount(XResultSet rs)
  {
    if (rs == null)
      return 0;

    int numCols = 0;
    try {
      XResultSetMetaDataSupplier rsMetaSupp = 
                          Lo.qi(XResultSetMetaDataSupplier.class, rs);
      XResultSetMetaData rsmd = rsMetaSupp.getMetaData();  
      numCols = rsmd.getColumnCount();
    }
    catch(SQLException e) {
      System.out.println(e);
    }
    return numCols;
  }  // end of getColumnCount()



  public static String[] getFieldNames(XResultSet rs)
  {
    if (rs == null) {
      System.out.println("No results set to print");
      return null;
    }

    String[] fieldNames = null;
    try {
      XResultSetMetaDataSupplier rsMetaSupp = 
                     Lo.qi(XResultSetMetaDataSupplier.class, rs);
      XResultSetMetaData rsmd = rsMetaSupp.getMetaData();
  
      int numCols = rsmd.getColumnCount();
      fieldNames = new String[numCols];
      for (int i = 0; i < numCols; i++)
        fieldNames[i] = rsmd.getColumnName(i+1);
    }
    catch(SQLException e) {
      System.out.println(e);
    }
    return fieldNames;
  }  // end of getFieldNames()



  public static void printResultSet(XResultSet rs)
  {
    if (rs == null) {
      System.out.println("No results set to print");
      return;
    }

    try {
      XResultSetMetaDataSupplier rsMetaSupp = 
               Lo.qi(XResultSetMetaDataSupplier.class, rs);
      XResultSetMetaData rsmd = rsMetaSupp.getMetaData();
      // printColTypes(rsmd);
  
      int tableWidth = 0;
      int numCols = rsmd.getColumnCount();

      String colName;
      // System.out.println();  
      for (int i = 0; i < numCols; i++) {
        if (i > 0) {
          System.out.print(", ");
          tableWidth += 2;
        }
        colName = rsmd.getColumnName(i+1);
        System.out.printf("%10s", colName);
        tableWidth += Math.max(colName.length(), 10);
      }
      System.out.println();

      for (int i=0; i < tableWidth; i++)
        System.out.print("-");
      System.out.println();

      XRow xRow = Lo.qi(XRow.class, rs);
      while (rs.next()) {
        for (int i = 0; i < numCols; i++) {
          if (i > 0) 
            System.out.print(", ");
          System.out.printf("%10s", xRow.getString(i+1) );
        }
        System.out.println();  
      }
      System.out.println();  
    }
    catch(SQLException e) {
      System.out.println(e);
    }
  }  // end of printResultSet()


  public static void displayResultSet(XResultSet rs)
  {  displayResultSet(rs, "Result set", 0);  }


  public static void displayResultSet(XResultSet rs, String title, int offset)
  {
    if (rs == null) {
      System.out.println("No results set to display");
      return;
    }

    try {
      XResultSetMetaDataSupplier rsMetaSupp = 
               Lo.qi(XResultSetMetaDataSupplier.class, rs);
      XResultSetMetaData rsmd = rsMetaSupp.getMetaData();

      // names of columns
      int numCols = rsmd.getColumnCount();
      String[] headers = new String[numCols];
      for (int i=0; i < numCols; i++) {
        headers[i] = rsmd.getColumnName(i+1);
        // System.out.println("header " + i + ": " + headers[i]);
      }
      // create table with column heads
      DefaultTableModel tableModel = new DefaultTableModel(headers, 0);
      JTable table = new JTable(tableModel);

      // fill table with XResultSet contents, one row at a time
      XRow xRow = Lo.qi(XRow.class, rs);
      while (rs.next())
        tableModel.addRow( getRow(xRow, numCols)); 

      // resize columns so data is visible
      table.setAutoResizeMode(JTable.AUTO_RESIZE_ALL_COLUMNS);
      resizeColumnWidth(table);
      
      SwingUtilities.invokeLater( new Runnable() {
        public void run() 
        {  JFrame frame = new JFrame();
           frame.setBounds(offset, offset, 400, 200);
           // frame.setSize(400,200);
           frame.setTitle(title);
           frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);  // so don't kill main
           frame.add(new JScrollPane(table), BorderLayout.CENTER);
           frame.setVisible(true); 
        }
      });
    }
    catch(SQLException e) {
      System.out.println(e);
    }
  }  // end of displayResultSet()


  private static void resizeColumnWidth(JTable table) 
  {
    TableColumnModel columnModel = table.getColumnModel();
    for (int col = 0; col < table.getColumnCount(); col++) {
      int width = 50;    // min width
      for (int row = 0; row < table.getRowCount(); row++) {
        TableCellRenderer renderer = table.getCellRenderer(row, col);
        Component comp = table.prepareRenderer(renderer, row, col);
        width = Math.max(comp.getPreferredSize().width, width);
      }
      columnModel.getColumn(col).setPreferredWidth(width);
    }
  }  // end of resizeColumnWidth()



   private static Object[] getRow(XRow xRow, int numCols) 
                                              throws SQLException
   // return current row of resultset as an array
   { 
     Object[] row = new Object[numCols];
     for (int i = 1; i <= numCols; i++) {
       row[i-1] = xRow.getString(i);
       // System.out.println("row " + (i-1) + ": " + row[i-1]);
     }
     return row;
   }  // end of getRow()




  public static Object[][] getResultSetArr(XResultSet rs)
  {
    if (rs == null) {
      System.out.println("No results set to convert");
      return null;
    }

    ArrayList<Object[]> rowsList = new ArrayList<Object[]>();
    int numCols = 0;
    try {
      XResultSetMetaDataSupplier rsMetaSupp = 
               Lo.qi(XResultSetMetaDataSupplier.class, rs);
      XResultSetMetaData rsmd = rsMetaSupp.getMetaData();
  
      numCols = rsmd.getColumnCount();
      Object[] rowVals = new Object[numCols];
      for (int i = 0; i < numCols; i++)
        rowVals[i] = rsmd.getColumnName(i+1);
      rowsList.add(rowVals);     // include the headers row 

      XRow xRow = Lo.qi(XRow.class, rs);
      while (rs.next()) {
        rowVals = new String[numCols];
        for (int i = 0; i < numCols; i++)
          rowVals[i] = xRow.getString(i+1);
        rowsList.add(rowVals);
      }
    }
    catch(SQLException e) {
      System.out.println(e);
    }

    int numRows = rowsList.size();
    if (numRows == 0) {
      System.out.println("The results set was empty");
      return null;
    }

    System.out.println("Result set array size: " + numRows + " x " + numCols);
    Object[][] rowsArr = new Object[numRows][numCols];
    for(int i=0; i < numRows; i++)
      rowsArr[i] = rowsList.get(i);
    return rowsArr;
  }  // end of getResultSetArr()



  public static void printResultSetArr(Object[][] rsa)
  {
    if (rsa == null) {
      System.out.println("No results set array to print");
      return;
    }

    // print header row 
    int tableWidth = 0;
    int numCols = rsa[0].length;
    String colName;
    // System.out.println();  
    for (int i = 0; i < numCols; i++) {
      if (i > 0) {
        System.out.print(", ");
        tableWidth += 2;
      }
      colName = (String)rsa[0][i];
      System.out.printf("%10s", colName);
      tableWidth += Math.max(colName.length(), 10);
    }
    System.out.println();

    // underline the header row
    for (int i=0; i < tableWidth; i++)
      System.out.print("-");
    System.out.println();

    // print the data rows
    int numRows = rsa.length;
    for(int j=1; j < numRows; j++) {  // start at row 1
      for (int i = 0; i < numCols; i++) {
        if (i > 0) 
          System.out.print(", ");
        System.out.printf("%10s", (String)rsa[j][i] );
      }
      System.out.println();  
    }
    System.out.println();  
  }  // end of printResultSetArr()



  // ----------- save a database as CSV files --------------------


  public static void saveDatabase(XConnection conn)
  {
    if (conn == null) {
      System.out.println("Connection is null");
      return;
    }

    ArrayList<String> tableNames = getTablesNames(conn);
    if (tableNames == null)
      System.out.println("No tables found in database");
    else {
      for(String tableName : tableNames)
        saveTable(tableName, conn);
    }
  }  // end of saveDatabase()


  public static void saveTable(String tableName, XConnection conn)
  {
    try { 
      XStatement statement = conn.createStatement();
      XResultSet rs = statement.executeQuery("SELECT * FROM " + tableName);
      System.out.println("Saving table: " + tableName);
      saveResultSet(rs, tableName + ".csv");
    }
    catch(com.sun.star.uno.Exception e) {
      System.out.println("Unable to save table: " + tableName);
      System.out.println(e);
    }
  }  // end of saveTable()




  public static void saveResultSet(XResultSet rs, String fnm)
  {
    System.out.println("  Writing result set to " + fnm);
    try {
      BufferedWriter bw = new BufferedWriter( new FileWriter(fnm));

      XResultSetMetaDataSupplier rsMetaSupp = 
               Lo.qi(XResultSetMetaDataSupplier.class, rs);
      XResultSetMetaData rsmd = rsMetaSupp.getMetaData();
      int numCols = rsmd.getColumnCount();

      // include the headers row 
      StringBuilder sb = new StringBuilder();
      for (int i = 0; i < numCols; i++) {
        if (i > 0) 
          sb.append(",");
        sb.append( rsmd.getColumnName(i+1) );
      }
      bw.write(sb.toString());
      bw.newLine();

      // add data rows
      XRow xRow = Lo.qi(XRow.class, rs);
      while (rs.next()) {
        sb = new StringBuilder();
        for (int i = 0; i < numCols; i++) {
          if (i > 0) 
            sb.append(",");
          sb.append(xRow.getString(i+1));
        }
        bw.write(sb.toString());
        bw.newLine();
      }    
      bw.close();
    } 
    catch (java.lang.Exception e) {
      System.out.println("Unable to store result set to " + fnm);
      System.out.println(e);
    }
  }  // end of saveResultSet();





  // ------------- read in SQL commands -------------------------------
 

  public static ArrayList<String> readCmds(String fnm)
  {
    ArrayList<String> cmds = new ArrayList<String>();
    BufferedReader br = null;
    try {
      br = new BufferedReader( new FileReader(fnm));
      System.out.println("Reading data from: " + fnm);

      String cmd = "";
      String line;
      boolean lineAdded = true;
      while ((line = br.readLine()) != null) {
        line = line.trim();
        //System.out.println("<" + line + ">");
        if ((line.length() == 0)  ||     // blank line or
            (line.startsWith("//"))) {   // comment line
          if (!lineAdded) {
            cmds.add(cmd);
            lineAdded = true;
          }
        }
        else if (line.startsWith("CREATE") ||    // start of new command line
                 line.startsWith("INSERT") || 
                 line.startsWith("UPDATE") ||
                 line.startsWith("DELETE") || 
                 line.startsWith("SELECT")) {
          if (!lineAdded)
            cmds.add(cmd);
          cmd = line;
          lineAdded = false;
        }
        else  // continuation of previous command
          cmd = cmd + " " + line;
      }
      if (!lineAdded)
        cmds.add(cmd);
    }
    catch (FileNotFoundException ex)
    {  System.out.println("Could not open: " + fnm); }
    catch (IOException ex) 
    {  System.out.println("Read error: " + ex); }
    finally {
      try {
        br.close();
      }
      catch (IOException ex) 
      {  System.out.println("Problem closing " + fnm); }
    }

    return cmds;
  }  // end of readCmds()



  // ---------- zip methods  for an ODB file -------------
  //  for examining content.xml and the database/ folder


  public static boolean isEmbedded(String fnm)
  {
    String embedFnm = getEmbeddedFnm(fnm);
    if (embedFnm == null)
      return false;
    return (embedFnm.equals(FIREBIRD_EMBEDDED) ||
            embedFnm.equals(HSQL_EMBEDDED));
  }


  public static boolean isFirebirdEmbedded(String fnm)
  {  
    String embedFnm = getEmbeddedFnm(fnm);
    if (embedFnm == null)
      return false;
    return (embedFnm.equals(FIREBIRD_EMBEDDED));  
  }


  public static boolean isHSQLEmbedded(String fnm)
  {  
    String embedFnm = getEmbeddedFnm(fnm);
    if (embedFnm == null)
      return false;
    return (embedFnm.equals(HSQL_EMBEDDED));
  }



  public static String getEmbeddedFnm(String fnm)
  /* return the name of the embedded database stored in fnm by looking
     inside its content.xml file */
  {
    FileIO.makeDirectory(TEMP_DIR);
    String contentFnm = TEMP_DIR + "content.xml";
    if (unzipContent(fnm, contentFnm)) {
      String embedRes = getEmbeddedResource(contentFnm);
     // if (embedRes == null)
     //   System.out.println("No embedded db info found");
     // else
     //   System.out.println("Found embedded db name: \"" + embedRes + "\"");
      FileIO.deleteFile(contentFnm);
      return embedRes;
    }
    else {
      System.out.println("Could not find content.xml inside " + fnm);
      return null;
    }
  }  // end of getEmbeddedFnm()




  public static boolean unzipContent(String fnm, String contentFnm)
  // unzip content.xml inside fnm, saving it to contentFnm
  {
    // System.out.println("Attempting to unzip file: " + fnm);
    boolean foundContent = false;
    try {
      ZipFile zipFile = new ZipFile(fnm);
      // search through file entries in the zipfile
      Enumeration zipEntries = zipFile.entries();
      while (zipEntries.hasMoreElements()) {
        ZipEntry zipEntry = (ZipEntry) zipEntries.nextElement();
        if (zipEntry.getName().startsWith("content.xml"))
          foundContent = Base.writeToFile(contentFnm, zipFile, zipEntry);
      }
      zipFile.close();
    }
    catch (java.lang.Exception e) {
      System.out.println(e);
    }
    return foundContent;
  } // end of unzipContent()



  public static boolean writeToFile(String fnm, ZipFile zipFile, ZipEntry zipEntry)
  // write the zipped data in zipEntry in zipFile into fnm
  {
    byte[] buffer = new byte[1024];
    try {
      InputStream in = zipFile.getInputStream(zipEntry);
      BufferedOutputStream out = new BufferedOutputStream( new FileOutputStream(fnm));
      int len;
      while ((len = in.read(buffer)) >= 0)
        out.write(buffer, 0, len);
      out.close();
      in.close(); 
      System.out.println("Created file: " + fnm);
      return true;
    }
    catch (java.lang.Exception e) {
      System.out.println("Could not create file: " + fnm);
      System.out.println(e);
      return false;
    }
  }  // end of writeToFile()




  private static String getEmbeddedResource(String fnm)
  /* A ODB file using an embedded database will have a 
     db:connection-resource xlink:href attribute stored in 
     its content.xml file, fnm
  */
  {
    DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
    factory.setIgnoringElementContentWhitespace(true);
    factory.setNamespaceAware(true);

    String resource = null;
    try {
      DocumentBuilder builder = factory.newDocumentBuilder();
      Document doc = builder.parse( new File(fnm));

      XPathFactory xpathFactory = XPathFactory.newInstance();
      XPath xpath = xpathFactory.newXPath();
      xpath.setNamespaceContext( new UniversalNamespaceResolver(doc) );

      XPathExpression expr =
                xpath.compile("//db:connection-resource/@xlink:href");
      resource = (String) expr.evaluate(doc, XPathConstants.STRING);
      if ((resource == null) || resource.equals("")) {
        System.out.println("Connection resource not found");
        return null;
      }
    }
    catch (XPathExpressionException e)
    {  System.out.println(e);  }
    catch (ParserConfigurationException e)
    {  System.out.println(e);  }
    catch (SAXException e)
    {  System.out.println(e);  }
    catch (IOException e) 
    {  System.out.println(e);  }
    return resource;
  }  // end of getEmbeddedResource()




  public static String getLinkedFnm(String fnm)
  /* return the name of the linked database named in fnm  by looking
     inside its content.xml file */
  {
    FileIO.makeDirectory(TEMP_DIR);
    String contentFnm = TEMP_DIR + "content.xml";
    if (unzipContent(fnm, contentFnm)) {
      String linkRes = getLinkedResource(contentFnm);
     // if (linkRes == null)
     //   System.out.println("No linked db info found");
     // else
     //   System.out.println("Found linked db name: \"" + linkRes + "\"");
      FileIO.deleteFile(contentFnm);
      return linkRes;
    }
    else {
      System.out.println("Could not find content.xml inside " + fnm);
      return null;
    }
  }  // end of getLinkedFnm()



  private static String getLinkedResource(String fnm)
  /* A ODB file that links to a database will have a 
     db:file-based-database xlink:href attribute stored in 
     its content.xml file, fnm
  */
  {
    DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
    factory.setIgnoringElementContentWhitespace(true);
    factory.setNamespaceAware(true);

    String resource = null;
    try {
      DocumentBuilder builder = factory.newDocumentBuilder();
      Document doc = builder.parse( new File(fnm));

      XPathFactory xpathFactory = XPathFactory.newInstance();
      XPath xpath = xpathFactory.newXPath();
      xpath.setNamespaceContext( new UniversalNamespaceResolver(doc) );

      XPathExpression expr =
                xpath.compile("//db:file-based-database/@xlink:href");
      resource = (String) expr.evaluate(doc, XPathConstants.STRING);
      if ((resource == null) || resource.equals("")) {
        System.out.println("database link not found");
        return null;
      }
    }
    catch (XPathExpressionException e)
    {  System.out.println(e);  }
    catch (ParserConfigurationException e)
    {  System.out.println(e);  }
    catch (SAXException e)
    {  System.out.println(e);  }
    catch (IOException e) 
    {  System.out.println(e);  }
    return resource;
  }  // end of getLinkedResource()



  // ----------- extract embedded DB files to TEMP_DIR  ---------------


  public static ArrayList<String> extractEmbedded(String fnm)
  {
    String embedFnm = getEmbeddedFnm(fnm);
    if (embedFnm == null) {
      System.out.println(fnm + " is not an embedded HSQL/Firebird database");
      return null;
    }
    else if (embedFnm.equals(HSQL_EMBEDDED)) {
      System.out.println(fnm + " is an embedded HSQL database");
      return unzipFiles(fnm, HSQLDB);
    }
    else if (embedFnm.equals(FIREBIRD_EMBEDDED)) {
      System.out.println(fnm + " is an embedded Firebird database");
      return unzipFiles(fnm, FIREBIRD);
      // fixOwnership(Base.FB_FNM + ".fdb");
    }
    else {
      System.out.println(fnm + " holds an unknown embedded database: " + embedFnm);
      return null;
    }
  }  // end of extractEmbedded()


  private static ArrayList<String> unzipFiles(String fnm, int embedType)
  {
    System.out.println("Unzipping " + fnm + " to " + TEMP_DIR);
    ArrayList<String> dbFnms = new ArrayList<String>();
    try {
      ZipFile zipFile = new ZipFile(fnm);
      Enumeration zipEntries = zipFile.entries();
      while (zipEntries.hasMoreElements()) {
        ZipEntry zipEntry = (ZipEntry) zipEntries.nextElement();

        // If the file is in the database directory, extract it
        if (zipEntry.getName().startsWith(ZIP_DIR_NM)) {
          String dbFnm = storeFile(zipFile, zipEntry, embedType);
          if (dbFnm != null)
            dbFnms.add(dbFnm);
        }
      }
      zipFile.close();
    }
    catch (java.lang.Exception e) {
      System.out.println(e);
    }
    return dbFnms;
  } // end of unzipFiles()


  private static String storeFile(ZipFile zipFile, ZipEntry zipEntry, int embedType)
  {
    String zipName = zipEntry.getName();
    // System.out.println("Extracting zipped file: " + zipName);

    int zipDirLen = ZIP_DIR_NM.length();    // length of "database/"
    String dbFnm = null;
    if (embedType == HSQLDB)
      dbFnm = TEMP_DIR + HSQL_FNM + "." + zipName.substring(zipDirLen);
    else  // Firebird 
      dbFnm = TEMP_DIR + zipName.substring(zipDirLen);
    // System.out.println("   Creating: " + dbFnm + "\n");
    if (writeToFile(dbFnm, zipFile, zipEntry))
      return dbFnm;
    else 
      return null;
  }  // end of storeFile()


/*
  public static void fixOwnership(String fnm)
  // fix Firebird's 2.5.5. ownership problem by calling fbfix.bat 
  // on fnm in TEMP_DIR
  {
    try {
      Runtime.getRuntime().exec("cmd /c fbfix.bat " + fnm);
    }
    catch (java.lang.Exception e) {
      System.out.println("Unable to fix Firebird ownership: " + e);
    }
  }  // end of fixOwnership()
*/


  // -------------------------- RowSet methods ----------------------------


  public static XRowSet rowSetQuery(String fnm, String query)
  {
    XRowSet xRowSet = Lo.createInstanceMCF(XRowSet.class, "com.sun.star.sdb.RowSet");

    Props.setProperty(xRowSet, "DataSourceName", FileIO.fnmToURL(fnm));
    Props.setProperty(xRowSet, "CommandType", CommandType.COMMAND);   // TABLE, QUERY or COMMAND
    Props.setProperty(xRowSet, "Command", query);
         // command could be a table, query name, or SQL, depending on the CommandType

    // if your database requires login
    // Props.setProperty(xRowSet, "User", "");
    // Props.setProperty(xRowSet, "Password", "");
    // most attributes are defined in sdbc Rowset 

    return xRowSet;
  }  // end of rowSetQuery()



  public static boolean canInsert(XRowSet xRowSet)
  {  return checkPrivilege(xRowSet, Privilege.INSERT);  }


  public static boolean canSelect(XRowSet xRowSet)
  {  return checkPrivilege(xRowSet, Privilege.SELECT);  }


  public static boolean canUpdate(XRowSet xRowSet)
  {  return checkPrivilege(xRowSet, Privilege.UPDATE);  }


  public static boolean canDelete(XRowSet xRowSet)
  {  return checkPrivilege(xRowSet, Privilege.DELETE);  }



  public static boolean checkPrivilege(XRowSet xRowSet, long privFlag)
  {
    Integer priv = (Integer) Props.getProperty(xRowSet, "Privileges");
    if (priv == null) {
      System.out.println("No privilege   information found");
      return false;
    }
    return ((priv.intValue() & privFlag) == privFlag);
  }  // end of checkPrivilege()



  // ---------------- database/table info using SDBCX ------------



  public static ArrayList<String> getTablesNames(XConnection conn)
  {
    XTablesSupplier tblsSupplier = Lo.qi(XTablesSupplier.class, conn);
    if (tblsSupplier == null) {
      System.out.println("No table supplier found");
      return null;
    }
    XNameAccess tables = tblsSupplier.getTables();
    String[] tableNms = tables.getElementNames();
    return new ArrayList<String>(Arrays.asList(tableNms));
  }  // end of getTablesNames()



  public static void displayTablesInfo(XConnection conn)
  {
    XTablesSupplier tblsSupplier = Lo.qi(XTablesSupplier.class, conn);
    XNameAccess tables = tblsSupplier.getTables();
    String[] tableNms = tables.getElementNames();
    System.out.println("\nNo. of tables: " + tableNms.length);
    for (int i = 0; i < tableNms.length; i++)
       displayTableInfo(conn, tableNms[i]);

    printGroups(conn);
  }  // end of displayTablesInfo()



  public static void displayTableInfo(XConnection conn, String tableNm)
  {
    System.out.println("Table: " + tableNm);
    XTablesSupplier tblsSupplier = Lo.qi(XTablesSupplier.class, conn);
    XNameAccess tables = tblsSupplier.getTables();
    try {
       displayTableProperties( tables.getByName(tableNm) );  // a sdbcx Table service
    }
    catch(com.sun.star.uno.Exception e)
    {  System.out.println("  No table found");  }
  }  // end of displayTablesInfo()



  private static void displayTableProperties(Object table) 
  // print attributes stored in sdbcx Table service
  {
    XPropertySet props = Lo.qi(XPropertySet.class, table);
    System.out.println("  Name: \"" + Props.getProperty(table, "Name") +"\"");
    System.out.println("  Catalog Name: \"" + Props.getProperty(table, "CatalogName") +"\"");
    System.out.println("  Schema Name: \"" + Props.getProperty(table, "SchemaName") +"\"");
    System.out.println("  Description: \"" + Props.getProperty(table, "Description") +"\"");

    // the following property is optional so we first must check if it exists
    if (Props.hasProperty(table, "Type"))
      System.out.println("  Type: \"" + Props.getProperty(table, "Type") +"\"");

    printColumns(table);
    printKeys(table);
    printIndexes(table);
    // printColumnsProperties(table);

    System.out.println();
  }  // end of displayTableProperties()




  private static void printColumns(Object table) 
  // print all columns of a XColumnsSupplier
  {
    XColumnsSupplier colsSupplier = Lo.qi(XColumnsSupplier.class, table);
    System.out.print("  Columns: ");
    XNameAccess columns = colsSupplier.getColumns();  // a sdbcx Column service
    String[] columnNames = columns.getElementNames();
    for (int i = 0; i < columnNames.length; i++)
      System.out.print("\"" + columnNames[i] + "\" ");
    System.out.println();
  }  // end of printColumns()



  private static void printKeys(Object table) 
  // print all keys inclusive the columns of a key
  // examine sdbcx Key service
  {
    XKeysSupplier keysSup = Lo.qi(XKeysSupplier.class, table);
    if (keysSup != null) {
      System.out.println("  Keys:");
      XIndexAccess keys = keysSup.getKeys();
      for (int i = 0; i < keys.getCount(); i++) {
        try {
          Object key = keys.getByIndex(i);
          if (key != null) {
            System.out.print("    " + (i+1) + ". " + Props.getProperty(key, "Name"));
            printKeyProperties(key);
            XColumnsSupplier keyColumnsSup = Lo.qi(XColumnsSupplier.class, key);
            printColumns(keyColumnsSup);
          }
          else
            System.out.println("    "  + (i+1) + ". No key");
        }
       catch(com.sun.star.uno.Exception e)
       {  System.out.println("    "  + (i+1) + ". No key with that index");  }
      }
    }
  }  // end of printKeys()



  private static void printKeyProperties(Object key)
  // examine sdbcx Key service properties
  {
    // System.out.print(" Name: " + Props.getProperty(key, "Name"));
    System.out.print(" Type: " + Props.getProperty(key, "Type"));
    System.out.print(" ReferencedTable: " + Props.getProperty(key, "ReferencedTable"));
    System.out.print(" UpdateRule: " + Props.getProperty(key, "UpdateRule"));
    System.out.print(" DeleteRule: " + Props.getProperty(key, "DeleteRule"));
    // System.out.println();
  }  // end of printKeyProperties()




  private static void printIndexes(Object table) 
  // print all keys inclusive the columns of a key
  {
    XIndexesSupplier xIndexesSup = Lo.qi(XIndexesSupplier.class, table);
    if (xIndexesSup != null) {
      System.out.println("  Indexes:");
      XNameAccess xIndexs = xIndexesSup.getIndexes();
      String[] indexNms = xIndexs.getElementNames();
      for (int i = 0; i < indexNms.length; i++) {
        System.out.print("    " + (i+1) + ". " + indexNms[i]);
        try {
          Object index = xIndexs.getByName(indexNms[i]);
          // printIndexProperties(index);
          XColumnsSupplier indexColumnsSup = Lo.qi(XColumnsSupplier.class, index);
          printColumns(indexColumnsSup);
        }
        catch(com.sun.star.uno.Exception e)
        {  System.out.println("    No index found");  }
      }
    }
  }  // end of printIndexes()



  private static void printColumnsProperties(Object table) 
  {
    XColumnsSupplier colsSupplier = Lo.qi(XColumnsSupplier.class, table);
    System.out.println("  Columns Props:");
    XNameAccess columns = colsSupplier.getColumns();
    String[] columnNames = columns.getElementNames();
    for (int i = 0; i < columnNames.length; i++) {
      // System.out.println("  \"" + columnNames[i] + "\"");
      try {
        Object column = columns.getByName(columnNames[i]);
        printColumnProperties(column);
      }
      catch(com.sun.star.uno.Exception e) {}
    }
    System.out.println();

  }  // end of printColumnsProperties()




  private static void printColumnProperties(Object column) 
  // examine sdbcx Column service properties
  {
    System.out.print("  -- Name: \"" + Props.getProperty(column, "Name") + "\"");
    System.out.print("  Type: " + Props.getProperty(column, "Type"));
    System.out.print("  TypeName: " + Props.getProperty(column, "TypeName"));
    System.out.print("  Precision: " + Props.getProperty(column, "Precision"));
    System.out.print("  Scale: " + Props.getProperty(column, "Scale"));
    System.out.print("  IsNullable: " + Props.getProperty(column, "IsNullable"));
    System.out.print("  IsAutoIncrement: " + Props.getProperty(column, "IsAutoIncrement"));
    System.out.print("  IsCurrency: " + Props.getProperty(column, "IsCurrency"));

    // the following properties are optional
    if (Props.hasProperty(column, "IsRowVersion"))
      System.out.print("  IsRowVersion: " + Props.getProperty(column, "IsRowVersion"));

    if (Props.hasProperty(column, "Description"))
      System.out.print("  Description: " + Props.getProperty(column, "Description"));

    if (Props.hasProperty(column, "DefaultValue"))
      System.out.print("  DefaultValue: " + Props.getProperty(column, "DefaultValue"));

    System.out.println();
  }  // end of printColumnProperties()




  private static void printIndexProperties(Object index)
  // examine sdbcx Index service properties
  {
    // System.out.print(" Name: " + Props.getProperty(index, "Name"));
    System.out.print(" Catalog: " + Props.getProperty(index, "Catalog"));
    System.out.print(" IsUnique: " + Props.getProperty(index, "IsUnique"));
    System.out.print(" IsPrimaryKeyIndex: " + Props.getProperty(index, "IsPrimaryKeyIndex"));
    System.out.print(" IsClustered: " + Props.getProperty(index, "IsClustered"));
    System.out.println();
  }  // end of printIndexProperties()




  public static void printGroups(XConnection conn) 
  // print all groups and the users with their privileges who belong to this group
  // examine sdbcx Group and User services properties
  {
    XGroupsSupplier groupsSupp = Lo.qi(XGroupsSupplier.class, conn);
    if (groupsSupp == null) {
      System.out.println("No group supplier found");
      return;
    }

    XNameAccess xGroups = groupsSupp.getGroups();
    if (xGroups == null) {
      System.out.println("No groups found");
      return;
    }

    System.out.println("--- Groups ---");
    String[] groupNms = xGroups.getElementNames();
    for (int i = 0; i < groupNms.length; i++) {
      System.out.println("    " + groupNms[i]);
      try {
        XUsersSupplier usersSupp = Lo.qi(
                             XUsersSupplier.class, xGroups.getByName(groupNms[i]));
        if (usersSupp != null) {
          XAuthorizable auth = Lo.qi(XAuthorizable.class, usersSupp);
          System.out.println("\tUsers:");
          XNameAccess xUsers = usersSupp.getUsers();
          String[] userNms = xUsers.getElementNames();
          for (int j = 0; j < userNms.length; j++)
            System.out.println("\t    " + userNms[j] + " Privileges: " +
                                auth.getPrivileges(userNms[j], PrivilegeObject.TABLE));
        }
      }
      catch(com.sun.star.uno.Exception e) {}
    }
  }  // end of printGroups()



  // ---------------- database/table info using SDBC ------------
  // utilizes select queries and metadata


  public static void reportDBInfo(XConnection conn)
  {
    try {
      XDatabaseMetaData md = conn.getMetaData();

      String productName = md.getDatabaseProductName();
      String productVersion = md.getDatabaseProductVersion();
      if ((productName == null) || productName.equals(""))
        System.out.println("No database info found");
      else
        System.out.println("DB:  " + productName + " v." + productVersion);

      String driverName = md.getDriverName();
      String driverVersion = md.getDriverVersion();
      if ((driverName == null) || driverName.equals(""))
        System.out.println("No driver info found");
      else
        System.out.println("SDBC driver:  " + driverName + " v." + driverVersion);
    } 
    catch (SQLException e) 
    {  System.out.println(e);  }
  }  // end of reportDBInfo()




  public static ArrayList<String> getTablesNamesMD(XConnection conn)
  // get table names using DatabaseMetaData
  {
    ArrayList<String> names = new ArrayList<String>();
    try {
      XDatabaseMetaData dm = conn.getMetaData();
      XResultSet rs = dm.getTables(null, null, "%", new String[]{"TABLE"});
      XRow xRow = Lo.qi(XRow.class, rs);
      while (rs.next())
        names.add(xRow.getString(3));    // 3 == table name
    }
    catch(SQLException e) {
      System.out.println(e);
    }
    return names;
  }  // end of getTablesNamesMD()




  public static void displayTablesSchema(XConnection conn)
  {  displayTablesSchema(conn, true);  }


  public static void displayTablesSchema(XConnection conn, boolean useSDBCX)
  {
    ArrayList<String> tableNames = (useSDBCX) ? getTablesNames(conn) :
                                                getTablesNamesMD(conn);
      // choose to use SDBCX os DatabaseMetaData
    if (tableNames == null) 
      System.out.println("No tables found in database");
    else {
      System.out.println("No. of tables: " + tableNames.size());
      ArrayList<String> columnNames;
      for(String tableName : tableNames) {
        System.out.print("  " + tableName + ":");
        columnNames = getColumnNames(conn, tableName);
        if (columnNames == null)
          System.out.println(" -- no column names --");
        else {
          for(String colName : columnNames)
            System.out.print(" \"" + colName + "\"");
          System.out.println("\n");
        }
      }
    }
  }  // end of displayTablesSchema()



  public static ArrayList<String> getColumnNames(XConnection conn, String tableName)
  {
    ArrayList<String> names = new ArrayList<String>();
    try {
      XDatabaseMetaData dm = conn.getMetaData();
      XResultSet rs = dm.getColumns(null, null, tableName, "%");
      XRow xRow = Lo.qi(XRow.class, rs);
      while (rs.next())
        names.add(xRow.getString(4));    // 4 == column name
    }
    catch(SQLException e) {
      System.out.println(e);
    }
    return names;
  }  // end of getColumnNames()



  public static void reportResultSetSupport(XConnection conn)
  {
    try {
      XDatabaseMetaData md = conn.getMetaData();
      System.out.println("Resultset Capabilities:");
      printTypeConcurrency(md, ResultSetType.FORWARD_ONLY, "forward only");
      printTypeConcurrency(md, ResultSetType.SCROLL_INSENSITIVE, "scrollable; db insensitive");
      printTypeConcurrency(md, ResultSetType.SCROLL_SENSITIVE, "scrollable; db sensitive");
    }
    catch(SQLException e) {
      System.out.println(e);
    }
  }  // end of reportResultSetSupport()


  private static void printTypeConcurrency(XDatabaseMetaData md, int t, String typeStr)
                                                                          throws SQLException
  /* Possible ResultSet Type Values: 
        FORWARD_ONLY: The result set is not scrollable (default)
        SCROLL_INSENSITIVE: The result set is scrollable but not sensitive to database changes. 
        SCROLL_SENSITIVE: The result set is scrollable and sensitive to database changes. 

     Possible 'Concurrency' Values: 
        READ_ONLY: The result set cannot be used to update the database (default)
        UPDATABLE: The result set can be used to update the database. 
  */
  {
    if (md.supportsResultSetType(t)) {
      System.out.print("  Supports " + typeStr);
      if (md.supportsResultSetConcurrency(t, ResultSetConcurrency.READ_ONLY))
        System.out.print(" + read-only");

      if (md.supportsResultSetConcurrency(t, ResultSetConcurrency.UPDATABLE))
        System.out.print(" + updatable");
      System.out.println();
    }
    //else
    //  System.out.println("  Does not support " + typeStr);
  }  // end of printTypeConcurrency()



  public static void reportFunctionSupport(XConnection conn)
  {
    try {
      XDatabaseMetaData md = conn.getMetaData();

      String fns = md.getNumericFunctions();
      System.out.println("Numeric functions:");
      System.out.println("  " + fns);

      fns = md.getStringFunctions();
      System.out.println("String functions:");
      System.out.println("  " + fns);

      fns = md.getSystemFunctions();
      System.out.println("System functions:");
      System.out.println("  " + fns);

      fns = md.getTimeDateFunctions();
      System.out.println("Time/Date functions:");
      System.out.println("  " + fns);
    }
    catch(SQLException e) {
      System.out.println(e);
    }
  }  // end of reportFunctionSupport()


  public static void reportSQLTypes(XConnection conn)
  {
    try {
      XDatabaseMetaData md = conn.getMetaData();
      XResultSet rs = md.getTypeInfo();
      System.out.println("\nSQL Type Info:");
      Base.displayResultSet(rs);
      // Base.printResultSet(rs);
    }
    catch(SQLException e) {
      System.out.println(e);
    }
  }  // end of reportSQLTypes()




  // ---------------- data source info ------------


  public static boolean isEmbedded(XOfficeDatabaseDocument dbDoc)
  {
    String embedFnm = getDataSourceType(dbDoc);
    if (embedFnm == null)
      return false;
    return (embedFnm.equals(FIREBIRD_EMBEDDED) ||
            embedFnm.equals(HSQL_EMBEDDED));
  }


  public static boolean isFirebirdEmbedded(XOfficeDatabaseDocument dbDoc)
  {  
    String embedFnm = getDataSourceType(dbDoc);
    if (embedFnm == null)
      return false;
    return (embedFnm.equals(FIREBIRD_EMBEDDED));  
  }


  public static boolean isHSQLEmbedded(XOfficeDatabaseDocument dbDoc)
  {  
    String embedFnm = getDataSourceType(dbDoc);
    if (embedFnm == null)
      return false;
    return (embedFnm.equals(HSQL_EMBEDDED));
  }



  public static String getDataSourceType(XOfficeDatabaseDocument dbDoc)
  {
    if (dbDoc == null) {
      System.out.println("Database is null");
      return null;
    }

    XDataSource dataSource = dbDoc.getDataSource();
    if (dataSource == null) {
      System.out.println("DataSource is null");
      return null;
    }
    else
      return (String) Props.getProperty(dataSource, "URL");
  }  // end of getDataSourceType()




  public static void printDataSourceInfo(XDataSource dataSource)
  {  Props.showObjProps("Data Source", dataSource);  }  


  public static boolean isPasswordRequired(XDataSource dataSource)
  {  return (Boolean) Props.getProperty(dataSource, "IsPasswordRequired"); } 


  public static boolean isReadOnly(XDataSource dataSource)
  {  return (Boolean) Props.getProperty(dataSource, "IsReadOnly"); } 



  public static void printRegisteredDataSources()
  // print all registered datasources
  {
    XDatabaseRegistrations dbRegs = 
                Lo.createInstanceMCF(XDatabaseRegistrations.class, 
                                       "com.sun.star.sdb.DatabaseContext");
    XNameAccess nmsAccess = Lo.createInstanceMCF(XNameAccess.class, 
                                       "com.sun.star.sdb.DatabaseContext");

    String dsNames[] = nmsAccess.getElementNames();
    System.out.println("Registered Data Sources (" + dsNames.length + ")");
    for (int i = 0; i < dsNames.length; ++i) {
      String dbLoc = null;
      try {
        dbLoc = dbRegs.getDatabaseLocation(dsNames[i]);
      }
      catch(com.sun.star.uno.Exception e) {}
      System.out.println("  " + dsNames[i] + " in " + dbLoc);
    }
    System.out.println();
  }  // end of printRegisteredDataSources()




  public static boolean isRegisteredDataSource(String dataSourceName)
  {
    XNameAccess nmsAccess = Lo.createInstanceMCF(XNameAccess.class, 
                                       "com.sun.star.sdb.DatabaseContext");

    String dsNames[] = nmsAccess.getElementNames();
    for (String dsn : dsNames)
       if (dsn.equals(dataSourceName))
         return true;
    return false;
  }  // end of isRegisteredDataSource()



  public static void registerDataSource(String dataSourceName, Object dataSource)
  {
    if (isRegisteredDataSource(dataSourceName))
      System.out.println("Data source name \"" + dataSourceName + 
                                                "\" already registered");
    else { // Register it with the database context
      try {
        XNamingService nmService = Lo.createInstanceMCF(XNamingService.class, 
                                       "com.sun.star.sdb.DatabaseContext");
        nmService.registerObject(dataSourceName, dataSource);
        System.out.println("Data source name \"" + dataSourceName + "\" registered");
      }
      catch(com.sun.star.uno.Exception e) {
        System.out.println("Unable to access data source naming service");
      }
    }
  }  // end of registerDataSource()



  public static Object createDataSource()
  {
    try {
      XSingleServiceFactory factory = Lo.createInstanceMCF(XSingleServiceFactory.class,
                                          "com.sun.star.sdb.DatabaseContext");
      return factory.createInstance();
    }
    catch (com.sun.star.uno.Exception e) {
      System.out.println("Error creating data source: " + e);
      return null;
    }
  }  // end of createDataSource()



  public static XDataSource getFileDataSource(String fnm)
  {  return getURLDataSource(FileIO.fnmToURL(fnm));  }


  public static XDataSource getURLDataSource(String fileURL)
  {
    try {
      XNameAccess nmsAccess = Lo.createInstanceMCF(XNameAccess.class, 
                                         "com.sun.star.sdb.DatabaseContext");
      XDataSource ds = Lo.qi(XDataSource.class, nmsAccess.getByName(fileURL) );
      if (ds == null) {
        System.out.println("No data source for " + fileURL);
        return null;
      }
      else {
        System.out.println("Found data source for " + fileURL);
        return ds;
      }
    }
    catch(com.sun.star.uno.Exception e) {
      System.out.println("Unable to access data source for " + fileURL +
                         ": " + e.getMessage());
      return null;
    }
  }  // end of getURLDataSource()



  // ----------- DriverManager and Drives info -----------------


  public static XDriverManager getDriverManager()
  { return Lo.createInstanceMCF(XDriverManager.class, 
                                 "com.sun.star.sdbc.DriverManager");
  }


  public static ArrayList<String> getSupportedDrivers()
  {
    XDriverManager dm = getDriverManager();
    XEnumerationAccess enumAccess = Lo.qi(XEnumerationAccess.class, dm);
    XEnumeration driversEnum = enumAccess.createEnumeration();
    if (driversEnum == null) {
      System.out.println("No drivers found");
      return null;
    }

    ArrayList<String> drivers = new ArrayList<String>();
    while(driversEnum.hasMoreElements()) {
      try {
        drivers.add( Info.getImplementationName(driversEnum.nextElement()) );
      }
      catch(com.sun.star.uno.Exception e) {}
    }
    return drivers;
  }  // end of getSupportedDrivers()



  public static XDriver getDriverByURL(String url)
  {
    //XDriverManager dm = getDriverManager();
    //XDriverAccess driverAccess = Lo.qi( XDriverAccess.class, dm);
    XDriverAccess driverAccess = Lo.createInstanceMCF(XDriverAccess.class, 
                                             "com.sun.star.sdbc.DriverManager");
    return driverAccess.getDriverByURL(url);
  }  // end of getDriverByURL()


  public static void printDriverProperties(XDriver driver, String url)
  {
    if (driver == null) {
      System.out.println("Driver is null");
      return;
    }
    try {
      System.out.println("Driver Name: " + Info.getImplementationName(driver));

      DriverPropertyInfo[] dpInfo = driver.getPropertyInfo(url, null);
      if (dpInfo == null) {
        System.out.println("Properties info for the driver is null");
        return;
      }
      System.out.println("No. of Driver properties: " + dpInfo.length);
      for(int i=0; i < dpInfo.length; i++)
         System.out.println("  " + dpInfo[i].Name + " = " + dpInfo[i].Value);
      System.out.println();
    }
    catch(SQLException e) {
      System.out.println("No properties info for the driver");
    }
  }  // end of printDriverProperties()



}  // end of Base class

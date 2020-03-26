
// Jdbc.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016

/* A growing collection of utility functions to make Java's JDBC
   easier to use. They are currently divided into the following
   groups:

   * connect to HSQL or Firebird database
   * execute methods: query, update
   * display info about the tables
   * display databases & tables
   * display result sets

  Utilities for using LO's Base are in Base.java
*/

package utils;

import java.awt.*;
import java.util.*;
import java.sql.*;

import javax.swing.*;
import javax.swing.table.*;



public class Jdbc
{

  // -----  connect to HSQL or Firebird database -----------------


  public static boolean isHSQLEmbedded(ArrayList<String> fnms)
  {  return hasName(Base.HSQL_FNM, fnms);  }


  public static boolean isFirebirdEmbedded(ArrayList<String> fnms)
  {  return hasName(Base.FB_FNM, fnms);  }


  private static boolean hasName(String nm, ArrayList<String> fnms)
  {
    for(String fnm : fnms)
      if (fnm.contains(nm))
        return true;
    return false;
  }  // end of hasName()




  public static Connection connectToDB(ArrayList<String> fnms)
  {
    if (isHSQLEmbedded(fnms))
      return connectToHSQL(Base.TEMP_DIR + Base.HSQL_FNM);
    else if (isFirebirdEmbedded(fnms))
      return connectToFB(Base.TEMP_DIR + Base.FB_FNM + ".fdb");
    else {
      System.out.println("Unrecognized embedded database");
      return null;
    }
  }  // end of connectToDB()



  public static Connection connectToHSQL(String filePath)
  {
   Connection conn = null;
    try {
      Class.forName("org.hsqldb.jdbcDriver");
      conn = DriverManager.getConnection("jdbc:hsqldb:file:" + filePath +
                              ";shutdown=true",  "SA", "");
                 // force database closure (shutdown) at connection close
                 // otherwise data, log and lock will not be deleted
    }
    catch (ClassNotFoundException e) {
      System.out.println("Failed to load JDBC-HSQLDB driver");
    }
    catch(SQLException e) {
      System.out.println(e);
    }
    return conn;
  }  // end of connectToHSQL()



  public static Connection connectToFB(String filePath)
  {
    Connection conn = null;
    try {
      Class.forName("org.firebirdsql.jdbc.FBDriver");    // requires Jaybird
      conn = DriverManager.getConnection(
                                     "jdbc:firebirdsql:embedded:" + filePath,
                                     "sysdba", "masterkey"); 
    }
    catch (ClassNotFoundException e) {
      System.out.println("Failed to load JDBC-Firebird driver");
    }
    catch(SQLException e) {
      System.out.println(e);
    }
    return conn;
  }  // end of connectToFB()




  // -------------------- execute methods: query, update -------------

  public static boolean exec(String cmd, Connection conn)
  {
    if (cmd.startsWith("CREATE"))
      return execute(cmd, conn);
    else if (cmd.startsWith("SELECT")) {
      ResultSet rs = executeQuery(cmd, conn);
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



  public static boolean execute(String stmtStr, Connection conn)
  {
    if (conn == null) {
      System.out.println("Connection is null");
      return false;
    }

    try { 
      Statement statement = conn.createStatement();
      statement.execute(stmtStr);
      System.out.println("Executed \"" + stmtStr + "\"");
      return true;
    }
    catch(SQLException e) {
      System.out.println("Unable to execute: \"" + stmtStr + "\"\n   " + e.getMessage());
      return false;
    }
  }  // end of execute()



  public static ResultSet executeQuery(String query, Connection conn)
  {
    if (conn == null) {
      System.out.println("Connection is null");
      return null;
    }

    try { 
      Statement statement = conn.createStatement();
      ResultSet rs = statement.executeQuery(query);
      System.out.println("Processed \"" + query + "\": " + (rs != null));
      return rs;
    }
    catch(java.lang.Exception e) {
      System.out.println("Unable to execute query: \"" + query + "\"\n   " + e.getMessage());
      return null;
    }
  }  // end of executeQuery()



  public static int executeUpdate(String query, Connection conn)
  {
    if (conn == null) {
      System.out.println("Connection is null");
      return -1;
    }

    try { 
      Statement statement = conn.createStatement();
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



  // ---------------- display info about the tables ------------
  // utilizes select queries and metadata



  public static void reportDBInfo(Connection conn)
  {
    try {
      DatabaseMetaData md = conn.getMetaData();

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



  public static void displayTablesInfo(Connection conn)
  {
    ArrayList<String> tableNames = getTablesNames(conn);
    if (tableNames == null) 
      System.out.println("No tables found in database");
    else {
      System.out.println("No. of tables: " + tableNames.size());
      for(String tableName : tableNames)
        displayTableInfo(tableName, conn);
    }
  }  // end of displayTablesInfo()


  public static ArrayList<String> getTablesNames(Connection conn)
  {
    ArrayList<String> names = new ArrayList<String>();
    try {
      DatabaseMetaData dm = conn.getMetaData();
      ResultSet rs = dm.getTables(null, null, "%", new String[]{"TABLE"});
      while (rs.next())
        names.add(rs.getString(3));    // 3 == table name
    }
    catch(SQLException e) {
      System.out.println(e);
    }
    return names;
  }  // end of getTablesNames()


  public static void displayTableInfo(String tableName, Connection conn)
  {
    ResultSet rs = executeQuery("SELECT * FROM \"" + tableName + "\"", conn);
    System.out.println("Table: \"" + tableName + "\" ==========");
    printResultSet(rs);
    //displayResultSet(rs);
  }  // end of displayTableInfo()





  public static void displayTablesSchema(Connection conn)
  {
    ArrayList<String> tableNames = getTablesNames(conn);
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



  public static ArrayList<String> getColumnNames(Connection conn, String tableName)
  {
    ArrayList<String> names = new ArrayList<String>();
    try {
      DatabaseMetaData dm = conn.getMetaData();
      ResultSet rs = dm.getColumns(null, null, tableName, "%");
      while (rs.next())
        names.add(rs.getString(4));    // 4 == column name
    }
    catch(SQLException e) {
      System.out.println(e);
    }
    return names;
  }  // end of getColumnNames()




  // ------------- display databases & tables ---------------------
  // as printed text (using DBTablePrinter class), as a JTable



  public static void printDatabase(Connection conn)
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
        DBTablePrinter.printTable(conn, tableName);
    }
  }  // end of saveDatabase()




  public static void displayDatabase(Connection conn)
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



  public static void displayTable(String tableName, Connection conn, int offset)
  {
    try { 
      Statement statement = conn.createStatement();
      ResultSet rs = statement.executeQuery("SELECT * FROM " + tableName);
      displayResultSet(rs, tableName, offset);
    }
    catch(SQLException e) {
      System.out.println(tableName + " display error: " + e);
    }
  }  // end of displayTable()



  // ------------- display result sets ---------------------
  // as printed text, as a JTable, as a 2D array


  public static void printResultSet(ResultSet rs)
  {
    if (rs == null) {
      System.out.println("No results set to print");
      return;
    }

    try {
      ResultSetMetaData rsmd = rs.getMetaData();
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

      while (rs.next()) {
        for (int i = 0; i < numCols; i++) {
          if (i > 0) 
            System.out.print(", ");
          System.out.printf("%10s", rs.getString(i+1) );
        }
        System.out.println();  
      }
      System.out.println();  
    }
    catch(SQLException e) {
      System.out.println(e);
    }
  }  // end of printResultSet()


  public static void displayResultSet(ResultSet rs)
  {  displayResultSet(rs, "Result set", 0);  }


  public static void displayResultSet(ResultSet rs, String title, int offset)
  {
    try {
      // if there are no records, display a message
      if (!rs.next()) {
        JOptionPane.showMessageDialog(null, title + "  contains no records");
        return;
      }
      ResultSetMetaData rsmd = rs.getMetaData();

      // names of columns
      int numCols = rsmd.getColumnCount();
      String[] headers = new String[numCols];
      for (int i = 0; i < numCols; i++) 
        headers[i] = rsmd.getColumnName(i+1);
      
      // create table with column heads
      DefaultTableModel tableModel = new DefaultTableModel(headers, 0);
      JTable table = new JTable(tableModel);
      
      // fill table with ResultSet contents, one row at a time
      do {
        tableModel.addRow( getRow(rs, numCols) ); 
      } while (rs.next());

      // resize columns so data is visible
      table.setAutoResizeMode(JTable.AUTO_RESIZE_ALL_COLUMNS);
      resizeColumnWidth(table);
      
      SwingUtilities.invokeLater( new Runnable() {
        public void run() 
        {  JFrame frame = new JFrame();
           frame.setBounds(offset, offset, 400, 200);
           // frame.setSize(400,200);
           frame.setTitle(title);
           frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
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



   private static Object[] getRow(ResultSet rs, int numCols) 
                                                      throws SQLException
   // return current row of resultset as an array
   { 
     Object[] row = new Object[numCols];
     for (int i = 1; i <= numCols; i++)
        row[i-1] = rs.getString(i);
     return row;
   }  // end of getRow()



  public static Object[][] getResultSetArr(ResultSet rs)
  {
    if (rs == null) {
      System.out.println("No results set to convert");
      return null;
    }

    ArrayList<Object[]> rowsList = new ArrayList<Object[]>();
    int numCols = 0;
    try {
      ResultSetMetaData rsmd = rs.getMetaData();
  
      numCols = rsmd.getColumnCount();
      Object[] rowVals = new Object[numCols];
      for (int i = 0; i < numCols; i++)
        rowVals[i] = rsmd.getColumnName(i+1);
      rowsList.add(rowVals);     // include the headers row 

      while (rs.next()) {
        rowVals = new String[numCols];
        for (int i = 0; i < numCols; i++)
          rowVals[i] = rs.getString(i+1);
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


}  // end of Jdbc class


// ViewState.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015


/* Store the view state information for a sheet. The details are
   extracted from a view state string for the sheet which is returned 
   as part of the string generated by XController.getViewData()

   The set methods should contain **MUCH** more error-checking.
   For example, changing the split mode should not leave 
   other fields holding inconsistent values.

   Based on a post by user Hanya to:
    https://forum.openoffice.org/en/forum/viewtopic.php?f=45&t=29195&p=133202&hilit=getViewData#p133202
*/

package utils;

public class ViewState
{
  // for moving the pane focus
  public static final int MOVE_UP = 0;
  public static final int MOVE_DOWN = 1;
  public static final int MOVE_LEFT = 2;
  public static final int MOVE_RIGHT = 3;


  private int cursorColumn, cursorRow;
  private int colSplitMode, rowSplitMode;
  private int verticalSplit, horizontalSplit;
  private int focusNum;
  private int columnLeftPane, columnRightPane;
  private int rowUpperPane, rowLowerPane;



  public ViewState(String state)
  /*  The state string has the format:
             0/4998/0/1/0/218/2/0/0/4988/4998
  */
  {
    String[] states = state.split("/");
    if (states.length != 11)
      System.out.println("Incorrect number of states");
    else {
      cursorColumn = parseInt(states[0]);   // 0: cursor position column
      cursorRow = parseInt(states[1]);      // 1: cursor position row

      colSplitMode = parseInt(states[2]);   // 2: column split mode
      rowSplitMode = parseInt(states[3]);   // 3: row split mode

      verticalSplit = parseInt(states[4]);  // 4: vertical split position
      horizontalSplit = parseInt(states[5]); // 5: horizontal split position

      focusNum = parseInt(states[6]);     // 6: focused pane number

      columnLeftPane = parseInt(states[7]);  // 7: left column index of left pane
      columnRightPane = parseInt(states[8]); // 8: left column index of right pane

      rowUpperPane = parseInt(states[9]);    // 9: top row index of upper pane
      rowLowerPane = parseInt(states[10]);   // 10: top row index of lower pane
    }
  }  // end of ViewState()


  public int parseInt(String s)
  {
    if (s == null)
      return 0;
    try {
      return Integer.parseInt(s);
    }
    catch (NumberFormatException ex){ 
      System.out.println(s + " could not be parsed as an int; using 0");
      return 0;
    }
  }  // end of parseInt()




  public int getCursorColumn()
  {  return cursorColumn;  }


  public void setCursorColumn(int colPos)
  { if (colPos < 0)
      System.out.println("Column position must be positive");
    else
      cursorColumn = colPos; 
  }


  public int getCursorRow()
  {  return cursorRow;  }


  public void setCursorRow(int rowPos)
  { if (rowPos < 0)
      System.out.println("Row position must be positive");
    else
      cursorRow = rowPos; 
  }


  public int getColumnSplitMode()
  {  return colSplitMode;  }


  public void setColumnSplitMode(boolean isSplit)
  {
    colSplitMode = ((isSplit) ? 1 : 0);
    if (colSplitMode == 0) {   // no column splitting
      verticalSplit = 0;
      if ((focusNum == 1) || (focusNum == 3))
        focusNum--;  // move focus to left
     }
   }


  public int getRowSplitMode()
  {  return rowSplitMode;  }

  public void setRowSplitMode(boolean isSplit)
  {  
    rowSplitMode = ((isSplit) ? 1 : 0); 
    if (rowSplitMode == 0) {   // no row splitting
      horizontalSplit = 0;
      if ((focusNum == 2) || (focusNum == 3))
        focusNum -= 2;  // move focus up
     }
  }



  public int getVerticalSplit()
  {  return verticalSplit;  }

  public void setVerticalSplit(int splitPos)
  { if (splitPos < 0)
      System.out.println("Position must be positive");
    else
      verticalSplit = splitPos;
  }


  public int getHorizontalSplit()
  {  return horizontalSplit;  }

  public void setHorizontalSplit(int splitPos)
  {  if (splitPos < 0)
      System.out.println("Position must be positive");
    else
      horizontalSplit = splitPos;
  }


  public int getPaneFocusNum()
  {  return focusNum;  }


  public void setPaneFocusNum(int n)
  {  
    if ((n < 0) || (n > 3)) {
      System.out.println("Focus number is out of range 0-3");
      return;
    }

    if ((horizontalSplit == 0) && ((n == 1) || (n == 3))) {
      System.out.println("No horizontal split, so focus number must be 0 or 2");
      return;
    }

    if ((verticalSplit == 0) && ((n == 2) || (n == 3))) {
      System.out.println("No horizontal split, so focus number must be 0 or 1");
      return;
    }
    focusNum = n; 
  }  // end of setPaneFocusNum()


  public void movePaneFocus(int dir)
  /*  The 4 posible view panes are numbered like so:
                        0  |  1
                        -------
                        2  |  3
      If there's no horizontal split then the panes are numbered 0 and 2.
      If there's no vertical split then the panes are numbered 0 and 1.
  */
  {  
    if (dir == MOVE_UP) {
      if (focusNum == 3)
        focusNum = 1;
      else if (focusNum == 2)
        focusNum = 0;
      else
        System.out.println("cannot move up");
    }
    else if (dir == MOVE_DOWN) {
      if (focusNum == 1)
        focusNum = 3;
      else if (focusNum == 0)
        focusNum = 2;
      else
        System.out.println("cannot move down");
    }
    else if (dir == MOVE_LEFT) {
      if (focusNum == 1)
        focusNum = 0;
      else if (focusNum == 3)
        focusNum = 2;
      else
        System.out.println("cannot move left");
    }
    else if (dir == MOVE_RIGHT) {
      if (focusNum == 0)
        focusNum = 1;
      else if (focusNum == 2)
        focusNum = 3;
      else
        System.out.println("cannot move right");
    }
    else
      System.out.println("Unknown move direction");
  }  // end of movePaneFocus()



  public int getColumnLeftPane()
  {  return columnLeftPane;  }

  public void setColumnLeftPane(int idx)
  { if (idx < 0)
      System.out.println("Index must be positive");
    else
      columnLeftPane = idx; 
  }


  public int getColumnRightPane()
  {  return columnRightPane;  }

  public void setColumnRightPane(int idx)
  { if (idx < 0)
      System.out.println("Index must be positive");
    else
      columnRightPane = idx; 
  }


  public int getRowUpperPane()
  {  return rowUpperPane;  }

  public void setRowUpperPane(int idx)
  { if (idx < 0)
      System.out.println("Index must be positive");
    else
      rowUpperPane = idx;
  }


  public int getRowLowerPane()
  {  return rowLowerPane;  }

  public void setRowLowerPane(int idx)
  { if (idx < 0)
      System.out.println("Index must be positive");
    else
      rowLowerPane = idx;
  }


  public void report()
  {
    System.out.println("Sheet View State");
    System.out.println("  Cursor pos (column, row): (" + cursorColumn +
                                                 ", " + cursorRow + ") or \"" +
                              Calc.getCellStr(cursorColumn, cursorRow) + "\"");

    if ((colSplitMode == 1) && (rowSplitMode == 1))
      System.out.println("  Sheet is split vertically and horizontally at " + 
                                verticalSplit + " / " + horizontalSplit);
    else if (colSplitMode == 1)
      System.out.println("  Sheet is split vertically at " + verticalSplit);
    else if (rowSplitMode == 1)
      System.out.println("  Sheet is split horizontally at " + horizontalSplit);
    else
      System.out.println("  Sheet is not split");

    System.out.println("  Number of focused pane: " + focusNum);

    System.out.println("  Left column indicies of left/right panes: " + 
                            columnLeftPane + " / " + columnRightPane);

    System.out.println("  Top row indicies of upper/lower panes: " + 
                            rowUpperPane + " / " + rowLowerPane);
    System.out.println();
  }  // end of report()



  public String toString()
  { return ( cursorColumn + "/" + cursorRow + "/" +
             colSplitMode + "/" + rowSplitMode + "/" + 
             verticalSplit + "/" + horizontalSplit + "/" + 
             focusNum + "/" + 
             columnLeftPane + "/" + columnRightPane + "/" + 
             rowUpperPane + "/" +  rowLowerPane  ); 
  }

}  // end of ViewState class
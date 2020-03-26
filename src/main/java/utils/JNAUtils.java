
// JNAUtils.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, March 2015

/*  Windows and window control utilities implemented using
    JNA v.4.1.0 (https://github.com/twall/jna).
    Docs at: http://twall.github.io/jna/4.1.0/

       * window-related
       * button-related
       * keys related
       * process-related
*/

package utils;

import java.awt.*;
import java.io.*;
import java.util.*;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;


import com.sun.jna.*;
import com.sun.jna.win32.*;
import com.sun.jna.ptr.*;

import com.sun.jna.platform.win32.*;
import com.sun.jna.platform.win32.WinDef.*;
import com.sun.jna.platform.win32.WinUser.*;


public class JNAUtils
{
  private static final String OFFICE_PROCESS = "soffice";
  private static final String OFFICE_CLASS_NAME = "SALFRAME";

  private static final int WAIT_TIME = 500;   // ms

  private static final int BUF_SIZE = 512;

  // used when setting window display in showWindow()
  // from http://msdn.microsoft.com/en-us/library/windows/desktop/ms633548(v=vs.85).aspx
  public static final int SW_HIDE = 0;   
  public static final int SW_SHOW = 5;
  public static final int SW_MAXIMIZE = 3; 
  public static final int SW_MINIMIZE = 6; 
  public static final int SW_RESTORE = 9;
  public static final int SW_ENABLE = 64;
  public static final int SW_DISABLE = 65;

  public static final int SW_SHOWNORMAL = 1;     
  public static final int SW_SHOWMINIMIZED = 2;     
  public static final int SW_SHOWMAXIMIZED = 3;  
  public static final int SW_SHOWNOACTIVATE = 4;
  public static final int SW_SHOWMINNOACTIVE = 7;
  public static final int SW_SHOWNA = 8;
  public static final int SW_SHOWDEFAULT = 10;
  public static final int SW_FORCEMINIMIZE = 11;



  public interface  User32Ext extends User32
  {
    // Create an extended JNA proxy for user32.dll ...
    User32Ext INSTANCE = (User32Ext) Native.loadLibrary("user32", User32Ext.class,
                                        W32APIOptions.DEFAULT_OPTIONS);

   // docs at http://twall.github.io/jna/3.4.0/javadoc/com/sun/jna/platform/win32/User32.html


    // ---------------- additional methods --------------
/*
    int CascadeWindows(HWND hwnd, int how, RECT rect, int kids, HWND KidHandles);
    boolean EnableWindow(HWND handle, boolean isEnabled);
    int GetDlgCtrlID(HWND hWnd);
    HWND GetDesktopWindow();
    HWND GetTopWindow(HWND hWnd);
    boolean IsWindow(HWND handle);
    boolean IsWindowEnabled(HWND handle);
    int TileWindows(HWND hwnd, int how, RECT rect, int kids, HWND kidsHandles);
    HWND WindowFromPoint(int xPoint, int yPoint);
*/
  }  // end of User32Ext interface

  public static User32Ext user32ext = User32Ext.INSTANCE;


  public static Kernel32 kernel32 = (Kernel32) Native.loadLibrary(Kernel32.class, 
                                                   W32APIOptions.UNICODE_OPTIONS);


  public interface Psapi extends StdCallLibrary 
  {
    Psapi INSTANCE = (Psapi) Native.loadLibrary("Psapi", Psapi.class, 
                                             W32APIOptions.UNICODE_OPTIONS);

    int GetModuleFileNameExA(WinNT.HANDLE hProcess, WinNT.HANDLE hModule, 
                                              byte[] name, int nSize);
  }  // end of Psapi interface


  // ---------------- window-related methods -----------------------


	public static HWND getHandle()
  { return user32ext.FindWindow(OFFICE_CLASS_NAME, null); 
                // use class name, 2nd arg is window title
	}



  public static HWND getTitledHandle()
  // use the document's window title bar to find its associated window handle
  {
    String title = GUI.getTitleBar();
    if (title == null)
      return null;
    else
      return findTitledWin(title);
  }  // end of getTitledHandle()




  public static HWND findTitledWin(final String title)
  // return the handle of the window whose title starts with title
  {
    final ArrayList<HWND> handles = new ArrayList<HWND>();

    user32ext.EnumWindows(new WNDENUMPROC() {
      public boolean callback(HWND hWnd, com.sun.jna.Pointer arg1)
      {
        String winTitle = getTitle(hWnd);
        // System.out.println("  \"" + winTitle + "\"");
        if (winTitle.isEmpty())
          return true;
        if (winTitle.startsWith(title)) {
          // System.out.println("Found: \"" + winTitle + "\"");
          handles.add(hWnd);
        }
        return true;
      }
    }, null);

    HWND handle = null;
    if (handles.size() > 1) {
      System.out.println("Using the first matching window");
      handle = handles.get(0);
    }
    else if (handles.size() == 1)
      handle = handles.get(0);
    else if (handles.size() == 0)
      System.out.println("No matching window found for \"" + title + "\"");

    return handle;
  }  // end of findTitledWin()



  public static void winWait(String title)
  {
    Lo.wait(WAIT_TIME);
    while(true) {     // could run forever
      HWND hWnd = findTitledWin(title);
      if (hWnd == null)
        return;
      Lo.wait(WAIT_TIME);
    }
  }  // end of winWait()



  public static String handleString(HWND hWnd)
  // returns the handle as a hexadecimal string
  { if (hWnd == null)
      return null;
    else
      return Long.toHexString(Pointer.nativeValue(hWnd.getPointer()));
  }




  public static String getTitle(HWND hWnd)
  // get the window title via its handle
  {
    char[] buf = new char[BUF_SIZE];
    user32ext.GetWindowText(hWnd, buf, BUF_SIZE);
    return (new String(buf)).trim();
  }  // end of getTitle()



  public static String getClassName(HWND hWnd)
  {
    char[] buf = new char[BUF_SIZE];
    user32ext.GetClassName(hWnd, buf, BUF_SIZE);
    return (new String(buf)).trim();
  }  // end of getClassName()



  public static boolean setForegroundWindow(HWND hWnd) 
  {  return user32ext.SetForegroundWindow(hWnd);  }



  public static boolean showWindow(HWND hWnd, int flag)
  // Shows, hides, minimizes, maximizes, or restores a window
  {  return user32ext.ShowWindow(hWnd, flag);  }


  public static boolean winShow(HWND hWnd)
  {  return showWindow(hWnd, SW_SHOW);  }

  public static boolean winHide(HWND hWnd)
  {  return showWindow(hWnd, SW_HIDE);  }

  public static boolean winMaximize(HWND hWnd)
  {  return showWindow(hWnd, SW_MAXIMIZE);  }

  public static boolean winMinimize(HWND hWnd)
  {  return showWindow(hWnd, SW_MINIMIZE);  }

  public static boolean winRestore(HWND hWnd)
  {  return showWindow(hWnd, SW_RESTORE);    }






  //  ------------ button-related methods ---------------------


  public static HWND findButton(HWND hWnd, final String buttonLabel)
  // get the handle for a named button inside the given window 
  {
    final ArrayList<HWND> handles = new ArrayList<HWND>();

    user32ext.EnumChildWindows(hWnd, new WNDENUMPROC() {
      public boolean callback(HWND hWndControl, com.sun.jna.Pointer arg1)
      {
        String className = getClassName(hWndControl);
        String label = getTitle(hWndControl);
        System.out.println("Label: " + label);
        if (className.contains("Button") && label.contains(buttonLabel)) {
          System.out.println("Found button: \"" + label + "\"");
          handles.add(hWndControl);
        }
        return true;
      }
    }, null);

    HWND handle = null;
    if (handles.size() > 1) {
      System.out.println("Using the first matching button");
      handle = handles.get(0);
    }
    else if (handles.size() == 1)
      handle = handles.get(0);
    return handle;
  }  // end of findButton()




  public static Point getClickPoint(HWND handle)
  // calculate a point position in the center of the handle's bounds
  {
    Rectangle bounds = getBounds(handle);
    if (bounds == null) {
      System.out.println("Bounding rectangle is null");
      return null;
    }

    int xCenter = bounds.x + bounds.width/2;
    int yCenter = bounds.y + bounds.height/2;
    // System.out.println("Click Point: (" + xCenter + ", " + yCenter + ")");
    return new Point(xCenter, yCenter);
  }  // end of getClickPoint()



  public static Point getOffsetPoint(Point p1, int xDist, int yDist)
  {  return new Point( p1.x + xDist, p1.y + yDist);  }



  public static Rectangle getBounds(HWND handle)
  {
    if (handle == null) {
      System.out.println("Handle is null");
      return null;
    }

    RECT rect = new RECT();
    user32ext.GetWindowRect(handle, rect);
    if (rect == null) {
      System.out.println("Could not find bounding rectangle");
      return null;
    }
  
    if ((rect.left == 0) && (rect.right == 0))  {
      System.out.println("Bounding rectangle as 0 volume");
      return null;
    }
    return new Rectangle(rect.left, rect.top,  (rect.right - rect.left),
                                               (rect.bottom - rect.top));
  }  // end of getBounds()



  public static void doClick(final Point clickPt)
  // click the mouse at the specified point; doesn't use JNA!
  {
    if (clickPt == null) {
      System.out.println("Click point is null");
      return;
    }

    EventQueue.invokeLater(new Runnable() {
      public void run() {
        try {
          Point oldPos = MouseInfo.getPointerInfo().getLocation();
          Robot r = new Robot();
          r.mouseMove(clickPt.x, clickPt.y);
          Lo.delay(300);
          r.mousePress(InputEvent.BUTTON1_MASK);
          Lo.delay(300);
          r.mouseRelease(InputEvent.BUTTON1_MASK);
          r.mouseMove(oldPos.x, oldPos.y);
          System.out.println("Click completed");
        }
        catch(AWTException e)
        {  System.out.println("Unable to carry out click: " + e); }
      }
    });
  }  // end of doClick()



  public static void doDrag(final Point clickPt, final Point releasePt)
  // drag the cursor between the two points; doesn't use JNA!
  {
    if (clickPt == null) {
      System.out.println("Click point is null");
      return;
    }
    if (releasePt == null) {
      System.out.println("Release point is null");
      return;
    }

    EventQueue.invokeLater(new Runnable() {
      public void run() {
        try {
          Point oldPos = MouseInfo.getPointerInfo().getLocation();
          Robot r = new Robot();
          r.mouseMove(clickPt.x, clickPt.y);
          Lo.delay(300);
          r.mousePress(InputEvent.BUTTON1_MASK);
          Lo.delay(300);
          r.mouseMove(releasePt.x, releasePt.y);
          Lo.delay(300);
          r.mouseRelease(InputEvent.BUTTON1_MASK);
          r.mouseMove(oldPos.x, oldPos.y);
          // System.out.println("Drag completed");
        }
        catch(AWTException e)
        {  System.out.println("Unable to carry out Drag: " + e); }
      }
    });
  }  // end of doDrag()



  // --------------- keys related ------------------------



  public static void shootWindow()
  // take a screenshot of the window in focus; no use made of JNA
  {
    EventQueue.invokeLater(new Runnable() {
      public void run() {
        try {
          Robot r = new Robot();
          r.keyPress(KeyEvent.VK_ALT);
          r.keyPress(KeyEvent.VK_PRINTSCREEN);
          r.keyRelease(KeyEvent.VK_ALT);
          System.out.println("Screenshot of window completed");
        }
        catch(AWTException e)
        {  System.out.println("Unable to carry out screenshot: " + e); }
      }
    });
  }  // end of shootWindow()


  // -------------- process-related ----------------------


  public static ArrayList<Integer> getPIDs(String nm)
  // gather all the process IDs that starts with nm
  {
    final ArrayList<Integer> pids = new ArrayList<Integer>();

    WinNT.HANDLE snapshot = kernel32.CreateToolhelp32Snapshot(Tlhelp32.TH32CS_SNAPPROCESS, 
                                                               new WinDef.DWORD(0));
    Tlhelp32.PROCESSENTRY32.ByReference processEntry = 
                                     new Tlhelp32.PROCESSENTRY32.ByReference();         
    try  {
      while (kernel32.Process32Next(snapshot, processEntry)) {
        String processFnm = Native.toString(processEntry.szExeFile);
        if (processFnm.startsWith(nm)) {
          int pid = processEntry.th32ProcessID.intValue();
          pids.add(pid);
          //int numThreads = processEntry.cntThreads.intValue();
          //System.out.println("Found " + nm + " (" + processFnm + ") with pid=" + pid +
          //                  "; num threads: " + numThreads);
        }
      }
    }
    finally {
      kernel32.CloseHandle(snapshot);
    }
    return pids;
  }  // end of getPIDs()




  public static boolean killOffice() 
  // or use Lo.kilOffice()
  {
    WinNT.HANDLE snapshot = kernel32.CreateToolhelp32Snapshot(Tlhelp32.TH32CS_SNAPPROCESS, 
                                                               new WinDef.DWORD(0));
    Tlhelp32.PROCESSENTRY32.ByReference processEntry = 
                                      new Tlhelp32.PROCESSENTRY32.ByReference();         
    boolean killedOffice = false;
    try  {
      while (kernel32.Process32Next(snapshot, processEntry)) {
        String processFnm = Native.toString(processEntry.szExeFile);
        if (processFnm.startsWith(OFFICE_PROCESS)) {
          int pid = processEntry.th32ProcessID.intValue();
          int numThreads = processEntry.cntThreads.intValue();
          System.out.println("Found Office with pid = " + pid + 
                                                     "; num threads: " + numThreads);
          WinNT.HANDLE handle = kernel32.OpenProcess ( 
                              0x0400| /* PROCESS_QUERY_INFORMATION */
                              0x0800| /* PROCESS_SUSPEND_RESUME */
                              0x0010|
                              0x0001| /* PROCESS_TERMINATE */
                              0x00100000 /* SYNCHRONIZE */,
                           false, pid);
          if (handle != null) {
            kernel32.TerminateProcess(handle, 0);
            killedOffice = true;
            String handleName = getHandleName(handle);
            if (handle != null)
              System.out.println("Office (" + handleName + ") killed");
            else
              System.out.println("Office killed");
          }
          else
            System.out.println("Could not construct Office process handle");
        }
      }
    }
    finally {
      kernel32.CloseHandle(snapshot);
    }

    if (!killedOffice)
      System.out.println("Office not found");
    return killedOffice;
  }  // end of killOffice()



  private static String getHandleName(WinNT.HANDLE handle)
  {
    byte[] name = new byte[1024];
    int len = Psapi.INSTANCE.GetModuleFileNameExA(handle, null, name, 1024);
    if (len == 0)
      return null;
    else
      return Native.toString(name);
  }  // end of getHandleName()



  public static ArrayList<HWND> getProcessHandles(int processID)
  // get the list of window handles owned by the given process
  {
    final ArrayList<HWND> handles = new ArrayList<HWND>();

    user32ext.EnumWindows(new WNDENUMPROC() {
      public boolean callback(HWND hWnd, com.sun.jna.Pointer arg1)
      {
        String winTitle = getTitle(hWnd);
        IntByReference iptr = new IntByReference();
        user32ext.GetWindowThreadProcessId(hWnd, iptr);
        int pid = iptr.getValue();

        if (pid == processID) {
          // System.out.println("Found: \"" + winTitle + "\"  for " + processID);
          handles.add(hWnd);
        }
        return true;
      }
    }, null);

    if (handles.size() == 0)
      System.out.println("No matching window handles found for " + processID);

    return handles;
  }  // end of getProcessHandles()



}  // end of JNAUtils class

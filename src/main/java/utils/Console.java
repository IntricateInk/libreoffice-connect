
// Console.java

// RJHM van den Bergh, rvdb@comweb.nl
// http://www.comweb.nl/java/Console/Console.html

/* Redirect all applications output sent through System.out and System.err 
   to a text area in a JFrame. 

   Useful for debugging Addons and macros executing inside Office.
*/

package utils;

import java.io.*;
import java.awt.*;
import java.awt.event.*;
import javax.swing.*;



public final class Console implements Runnable
{
  // singleton pattern
  // private static volatile Console instance;


  private JFrame frame;
  private JTextArea textArea;
  private boolean isStandAlone;


  private final PipedInputStream pin1 = new PipedInputStream();    // for stdout
  private final PipedInputStream pin2 = new PipedInputStream();    // for stderr

  private Thread reader1;   // for pin1
  private Thread reader2;   // for pin2
  private volatile boolean hasFinished;  // signals the threads that they should exit

/*
  public static Console getInstance() 
  { return Console.getInstance(false);  }


  public static Console getInstance(boolean isAlone) 
  // get a single instance of this class
  {
    if (instance == null) {
      synchronized (Console.class) {
        if (instance == null)
          instance = new Console(isAlone);
      }
    }
    return instance;
  }  // end of getInstance()
*/


  public Console()
  {  this(false);  }     // called from another window/application


  public Console(boolean isAlone)
  {
    isStandAlone = isAlone;
    // create all components and add them
    frame = new JFrame("Output Console");
    frame.setAutoRequestFocus(false);   // since java 1.7
    Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();

    int x = (int) (screenSize.width*4/5);
    int y = 0;  // (int) (screenSize.height/2);
    int width = (int) (screenSize.width/5);
    int height = (int) (screenSize.height*3/4);
    frame.setBounds(x, y, width, height);

    // GUI: a text area and a "Clear" button
    textArea = new JTextArea();
    textArea.setEditable(false);
    textArea.setLineWrap(true);

    Font font = new Font("SANS_SERIF", Font.PLAIN, 12);
    textArea.setFont(font);

    JButton button = new JButton("clear");
    button.addActionListener( new ActionListener() {
      public void actionPerformed(ActionEvent e) 
      {  textArea.setText(""); }    // clears the text area
    });


    Container c = frame.getContentPane();
    c.setLayout(new BorderLayout());
    c.add(new JScrollPane(textArea), BorderLayout.CENTER);
    c.add(button, BorderLayout.SOUTH);

    //frame.setVisible(true);
    frame.addWindowListener(new WindowAdapter() {
      public void windowClosing(WindowEvent e) 
      { closeDown();
        if (isStandAlone)
          System.exit(0);
      }
    });

    // link stdout to pin 1
    try {
      PipedOutputStream pout1 = new PipedOutputStream(this.pin1);
      System.setOut(new PrintStream(pout1, true));
    }
    catch (IOException io) {
      textArea.append("Couldn't redirect STDOUT to this console\n" + io.getMessage());
    }
    catch (SecurityException se) {
      textArea.append("Couldn't redirect STDOUT to this console\n" + se.getMessage());
    }

    // link stderr to pin 2
    try {
      PipedOutputStream pout2 = new PipedOutputStream(this.pin2);
      System.setErr(new PrintStream(pout2, true));
    }
    catch (IOException io) {
      textArea.append("Couldn't redirect STDERR to this console\n" + io.getMessage());
    }
    catch (SecurityException se) {
      textArea.append("Couldn't redirect STDERR to this console\n" + se.getMessage());
    }

    hasFinished = false;

    // Starting two seperate threads to read from the PipedInputStreams
    reader1 = new Thread(this);   // for pin 1 reading
    reader1.setDaemon(true);
    reader1.start();

    reader2 = new Thread(this);   // for pin 2 reading
    reader2.setDaemon(true);
    reader2.start();

    System.out.println("Hello Console...");
  }  // end of Console()



  public void setVisible(boolean b)
  {  frame.setVisible(b);  }



  public synchronized void closeDown()
  {
    hasFinished = true;
    frame.setVisible(false);

    this.notifyAll(); // stop all threads
    try {
      reader1.join(500);
      pin1.close();
    }
    catch (Exception e) {}
    try {
      reader2.join(500);
      pin2.close();
    }
    catch (Exception e) {}

    frame.dispose();
  }  // end of closeDown()



  public synchronized void run()
  {
    try {
      while (Thread.currentThread() == reader1) {
        try {
          this.wait(100);
        }
        catch (InterruptedException ie) {}
        if (pin1.available() != 0) {
          String input = this.readLine(pin1);
          updateTextArea(input);
        }
        if (hasFinished) 
          return;
      }

      while (Thread.currentThread() == reader2) {
        try {
          this.wait(100);
        }
        catch (InterruptedException ie) {}
        if (pin2.available() != 0) {
          String input = this.readLine(pin2);
          updateTextArea(input);
        }
        if (hasFinished) 
          return;
      }
    }
    catch (Exception e) {
      updateTextArea("\nConsole reports an Internal error.\nThe error is: " + e);
    }
  }  // end of run()



  private void updateTextArea(String msg)
  {
    textArea.append(msg);
    int len = textArea.getDocument().getLength();
    textArea.setCaretPosition(len);
    textArea.repaint();
  }  // end of updateTextArea()



  public synchronized String readLine(PipedInputStream in) throws IOException
  {
    String input = "";
    do {
      int available = in.available();
      if (available == 0) 
        break;
      byte buf[] = new byte[available];
      in.read(buf);
      input = input + new String(buf, 0, buf.length);
    } while (!input.endsWith("\n") && !input.endsWith("\r\n") && !hasFinished);
    return input;
  }  // end of readLine()


  // ------------------------- test rig -----------------------

  public static void main(String[] arg)
  {  Console c = new Console(true);  
     c.setVisible(true);
  }

}  // end of Console class

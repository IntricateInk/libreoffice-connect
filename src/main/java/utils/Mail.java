
// Mail.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* E-mail utility functions for Office. 
   They are currently divided into the following groups:

     * sending e-mail using the MailServiceProvider service 

     * sending e-mail using the SimpleSystemMail (Windows)
       or SimpleCommandMail (Linux/Mac) service 

     * merge tasks for creating files, sending files to the printer,
       or sending files out as e-mail

   An example of Zawinski's Law :)
      http://www.catb.org/jargon/html/Z/Zawinskis-Law.html
*/

package utils;

import com.sun.star.beans.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;
import com.sun.star.frame.*;

import com.sun.star.mail.*;
import com.sun.star.system.*;
import com.sun.star.task.XJob; 

import com.sun.star.text.*;
import com.sun.star.sdb.*;



public class Mail
{
  // ----------------- MailServiceProvider service ----------------


  public static void sendEmail(String mailhost, int port,
                  String user, String password,
                  String to, String subject, String body, String fnm)
  // send an e-mail using the MailServiceProvider service
  /* Uses mailmerge.py in <OFFICE>\program
      -- error: "No SSL support included in this Python"
      -- related to ssleay32.dll
    
     Problems with conflicting versions of
     ssleay32.dll, libeay32.dll in Windows for OpenSSL support.
    
     Copy these two DLLs from the Libreoffice <OFFICE>\program\ dir to:
     <OFFICE>\program\python-core-????\lib\
    
    See bug reports:
      https://bugs.documentfoundation.org/show_bug.cgi?id=77354
  */
  {
    try {
      // create SMTP service
      XMailServiceProvider msp =
          Lo.createInstanceMCF(XMailServiceProvider.class,
                               "com.sun.star.mail.MailServiceProvider");
      if (msp == null) {
        System.out.println("Could not create MailServiceProvider");
        return;
      }

      XMailService service = msp.create(MailServiceType.SMTP);
      if (service == null) {
        System.out.println("Could not create SMTP MailService");
        return;
      }

      // setup service listener
      service.addConnectionListener( new XConnectionListener() {
        public void connected(EventObject e)
        {  System.out.println("  Connected to server " + getServerName(e)); }

        public void disconnected(EventObject e)
        {  System.out.println("  Disconnected"); }

        public void disposing(EventObject e) {}
      });


      // initialize service data: context and authenticator
      XCurrentContext xcc = new XCurrentContext() {
        public Object getValueByName(String name)
        {
          if (name.equals("ServerName"))
            return (Object) mailhost;
          else if (name.equals("Port"))
            return (Object) new Integer(port);
          else if (name.equals("ConnectionType"))
            return (Object) "Ssl";     // or "Insecure";
          else if (name.equals("Timeout"))
            return (Object)  new Integer(60);
          System.out.println("Do not recognize \"" + name + "\"");
          return null;
        }
      };

      XAuthenticator auth = new XAuthenticator() {
        public String getUserName()
        { return user;  }
        public String getPassword()
        { return password;  }
      };

      // connect to service
      service.connect(xcc, auth);
      // System.out.println("Isconnected: " + service.isConnected());

      // create message
      String from = user + "@" + mailhost;    // person sending this e-mail
      XMailMessage msg = com.sun.star.mail.MailMessage.create(
                 Lo.getContext(), to, from,
                 subject, new TextTransferable(body));

      if (fnm != null)
        msg.addAttachment( new MailAttachment(
                                new FileTransferable(fnm), fnm));

      // send message
      XSmtpService smtpService = Lo.qi(XSmtpService.class, service);
      System.out.println("  Sending e-mail...");
      smtpService.sendMailMessage(msg);

      service.disconnect();
    }
    catch(com.sun.star.uno.Exception e)
    {  System.out.println(e);  }

  }  // end of sendEmail()



  private static String getServerName(EventObject e)
  {
    XMailService service = (XMailService) e.Source;
    try {
      XCurrentContext xcc = service.getCurrentConnectionContext();
      return (String) xcc.getValueByName("ServerName");
    }
    catch(com.sun.star.io.NotConnectedException nce)
    {  System.out.println(nce);  
       return "??";
    }
  }  // end of getServerName()



  // ----------------- SimpleSystemMail service ----------------
  //                or SimpleCommandMail


  public static void sendEmailByClient(String to, String subject, 
                                        String body, String fnm) 
  // Send an e-mail using the SimpleSystemMail service (or SimpleCommandMail) to 
  // communicate with the OSes default e-mail client.
  // A "Confirm" dialog appears before the e-mail is sent.
  {
    System.out.println("Sending e-mail by client...");
    try {
      XSimpleMailClientSupplier mcSupp = 
          Lo.createInstanceMCF(XSimpleMailClientSupplier.class, 
                                   "com.sun.star.system.SimpleSystemMail");
                                   // windows e-mail client service
      if (mcSupp == null) {
        mcSupp = Lo.createInstanceMCF(XSimpleMailClientSupplier.class, 
                                   "com.sun.star.system.SimpleCommandMail");
                                    // returns null on Windows; used on Linux
        if (mcSupp == null) {
          System.out.println("Unable to create client using the SimpleSystemMail or SimpleCommandMail service");
          return;
        }
      }

      XSimpleMailClient mc = mcSupp.querySimpleMailClient();
            // defaults to ThunderBird on my system

      XSimpleMailMessage msg = mc.createSimpleMailMessage();
      msg.setRecipient(to);
      msg.setSubject(subject);

      XSimpleMailMessage2 msg2 = Lo.qi(XSimpleMailMessage2.class, msg);
      msg2.setBody(body);

      if (fnm != null) {
        String[] attachs = new String[1]; 
        attachs[0] = FileIO.getAbsolutePath(fnm);      // attachment
        msg.setAttachement(attachs);
      }

      mc.sendSimpleMailMessage(msg, SimpleMailClientFlags.NO_USER_INTERFACE);
           // hides GUI but still displays a "Confirm" dialog 
    }
    catch(com.sun.star.uno.Exception e)
    {  System.out.println(e);  }

  }  // end of sendEmailByClient()



  // ----------------------- mail merge task ---------------------

  /* Before mergeTask() can run:

       * A data spreadsheet must be connected to a same-named database
          which is a data source called <dataSourceName>. 
       
       * The spreadsheet sheet must be called <tableName>.

       * For e-mail, the spreadsheet must have a column called "E-mail"
         which holds e-mail addresses.
       
       * The merge fields must be copied from the database into the Writer
         template file, <templateFnm>.
        
     For info on how to do this read Chapter 11 of Writer Guide, or
     my "Sending E-mail" chapter.

     Office reports a crash at the end of mergeTask(), but terminates correctly.
  */


  public static void mergeLetter(String dataSourceName, String tableName,
                               String templateFnm, boolean isSingle)
  { 
    System.out.println("Merging letters to files...");
    mergeTask(dataSourceName, tableName, templateFnm, MailMergeType.FILE,
              isSingle, null, false,
              null, null, null);
  }  // end of mergeLetter()



  public static void mergePrint(String dataSourceName, String tableName,
                               String templateFnm, 
                               String printerName, boolean isMultipleJobs)
  { 
    System.out.println("Merging letters for printing...");
    mergeTask(dataSourceName, tableName, templateFnm, MailMergeType.PRINTER,
              false, printerName, isMultipleJobs, 
              null, null, null);
  }  // end of mergePrint()



  public static void mergeEmail(String dataSourceName, String tableName,
                               String templateFnm, 
                               String passwd, String subject, String body)
  { 
    System.out.println("Merging letters for sending as e-mail...");

    boolean isConfigured = checkMailConfig(passwd);
    System.out.println("--> Mailhost is " + 
              (isConfigured ? "" : "NOT ") + "configured");
    if (isConfigured)
      mergeTask(dataSourceName, tableName, templateFnm, MailMergeType.MAIL,
              false, null, false,
              passwd, subject, body);
  }  // end of mergeEmail()



  public static boolean checkMailConfig(String passwd)
  /* check that Office has e-mail settings for the mailhost
     ("MailServer"), the server's port, host username, and password
  */
  {
    boolean isConfigured = true;    // assume the best
    String serverName = Info.getRegItemProp("Writer/MailMergeWizard", "MailServer");
    if (serverName == null) {
       System.out.println(">> No mailhost name found; add one in Office");
       isConfigured = false;
    }
    else
      System.out.println(">> mailhost: " + serverName);


    String portStr = Info.getRegItemProp("Writer/MailMergeWizard", "MailPort");
    if (portStr == null) {
       System.out.println(">> No mailhost port found; add one in Office");
       isConfigured = false;
    }
    else
      System.out.println(">> mail port: " + portStr);


    String userName = Info.getRegItemProp("Writer/MailMergeWizard", "MailUserName");
    if (userName == null) {
       System.out.println(">> No mailhost username found; add one in Office");
       isConfigured = false;
    }
    else
      System.out.println(">> mail username: " + userName);


    String password = Info.getRegItemProp("Writer/MailMergeWizard", "MailPassword");
    if (password != null)
      System.out.println(">> Mailhost password found in Office; delete it for safety");
    else {
      if (passwd == null) {
        System.out.println(">> No mailhost password found; supply one at run-time");
        isConfigured = false;
      }
    }
    return isConfigured;
  }  // end of checkMailConfig()



  public static void mergeTask(String dataSourceName, String tableName,
                               String templateFnm, short outputType,
                   boolean isSingle,                             // for FILE
                   String printerName, boolean isMultipleJobs,   // for PRINTER
                   String passwd, String subject, String body)   // for MAIL
  { 
    XJob job = Lo.createInstanceMCF(XJob.class, "com.sun.star.text.MailMerge");
    if (job == null) {
      System.out.println("Could not create MailMerge service");
      return;
    }

    XPropertySet props = Lo.qi(XPropertySet.class, job); 

    // standard task properties
    Props.setProperty(props, "DataSourceName", dataSourceName);  
    Props.setProperty(props, "Command", tableName);
    Props.setProperty(props, "CommandType", CommandType.TABLE); 
    Props.setProperty(props, "DocumentURL", FileIO.fnmToURL(templateFnm)); 

    // vary properties based on output type
    Props.setProperty(props, "OutputType", outputType);
    if (outputType == MailMergeType.FILE) {
      Props.setProperty(props, "SaveAsSingleFile", isSingle);   
      Props.setProperty(props, "FileNamePrefix", "letter");  // hardwired
    }
    else if (outputType == MailMergeType.PRINTER) {
      Props.setProperty(props, "SinglePrintJobs", isMultipleJobs);
                                   // true means one print job for each letter
      PropertyValue[] pProps = 
         Props.makeProps("PrinterName", printerName, "Wait", true);  //  synchronous
             // from com.sun.star.view.PrintOptions
      Props.setProperty(props, "PrintOptions", pProps);
    }
    else if (outputType == MailMergeType.MAIL) {
      if (passwd != null)
        Props.setProperty(props, "OutServerPassword", passwd);

      Props.setProperty(props, "AddressFromColumn", "E-mail");  // hardwired column name
      Props.setProperty(props, "Subject", subject);
      Props.setProperty(props, "MailBody", body);
      
      Props.setProperty(props, "SendAsAttachment", true);
      Props.setProperty(props, "AttachmentName", "letter.pdf");  // hardwired filename
      Props.setProperty(props, "AttachmentFilter", "writer_pdf_Export");
    }

    // monitor task's execution
    XMailMergeBroadcaster xmmb = Lo.qi(XMailMergeBroadcaster.class, job);
    xmmb.addMailMergeEventListener( new XMailMergeListener() 
    {
      int count = 0;
      long start = System.currentTimeMillis();

      public void notifyMailMergeEvent(MailMergeEvent e)
      { count++;
        XModel model = e.Model;
        // Props.showProps("Mail merge event", model.getArgs());
        long currTime = System.currentTimeMillis();
        System.out.println("  Letter " + count + ": " +
                   (currTime - start) + "ms");
        start = currTime;
      }
    });

    try {
      job.execute(new NamedValue[0]); 
    }
    catch (com.sun.star.uno.Exception e) { 
      System.out.println("Could not start executing task: " + e); 
    } 
  }  // end of mergeTask()



}  // end of Mail class
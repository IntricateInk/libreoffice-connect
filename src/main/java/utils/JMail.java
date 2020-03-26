
// JMail.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Support functions for sending e-mail using Java alone, 
   without the Office API.

   sendEmail() uses JavaMail (javax.mail.jar), which is not
   a standard part of the JDK. It must be downloaded separately.

   Details on JavaMail: https://java.net/projects/javamail/pages/Home
     - JAR, samples, docs
   API docs: https://javamail.java.net/nonav/docs/api/
   FAQ: http://www.oracle.com/technetwork/java/javamail/faq/

   -------

   Office e-mail support functions are in Mail.java in the Utils/ folder.
*/

package utils;

import java.io.*;
import java.util.*;
import java.net.*;

import javax.mail.*;
import javax.mail.event.*;
import javax.mail.internet.*;
import javax.activation.*;

import com.sun.mail.smtp.*;

import java.awt.Desktop;


public class JMail
{

  // --------------------- JavaMail ----------------------------


  public static void sendEmail(String mailhost, int port, 
          String user, String password,
          String to, String subject, String body, String attachFnm) 
  {
    // store non-standard mail properties for the session
    Properties props = new Properties();
    props.put("mail.smtp.starttls.enable", "true");
    props.put("mail.smtp.ssl.trust", "*");  // no certificate needed
    props.put("mail.smtp.timeout", "60000");
    // props.put("mail.debug","true");

    try {
      Session session = Session.getInstance(props);

      // create message for the session
      SMTPMessage msg = new SMTPMessage(session);
      msg.setReturnOption(SMTPMessage.RETURN_HDRS);
      msg.setNotifyOptions(SMTPMessage.NOTIFY_SUCCESS |SMTPMessage.NOTIFY_FAILURE);

      msg.setFrom();  // uses default
      msg.setRecipient(Message.RecipientType.TO, new InternetAddress(to));
      msg.setSentDate(new Date());
    
      msg.setSubject(subject);
    
      if (attachFnm == null)
        msg.setText(body);
      else {   // add body text and file as attachments
        MimeBodyPart p1 = new MimeBodyPart();
        p1.setText(body);

        String mimeType = Info.getMIMEType(attachFnm);
        // System.out.println(mimeType);

        MimeBodyPart p2 = new MimeBodyPart();
        FileDataSource fds = new FileDataSource(attachFnm);
        p2.setDataHandler(new DataHandler(fds));  // add data,
        p2.setFileName(fds.getName());            // filename,
        p2.setHeader("Content-Type", mimeType);   // file's MIME type

        // create multipart
        Multipart mp = new MimeMultipart();
        mp.addBodyPart(p1);   // for body text 
        mp.addBodyPart(p2);   // for the attached file 
        msg.setContent(mp);
      }

      System.out.println("Sending e-mail using javax.mail...");

      // use SMTP with a user ID and password
      URLName url = new URLName("smtp", mailhost, port, "", user, password);
      Transport transport = new SMTPTransport(session, url);

      transport.addConnectionListener( new ConnectionListener() {
          public void opened(ConnectionEvent e) 
          {  System.out.println("  Connection opened to: " + e.getSource()); }
          
          public void disconnected(ConnectionEvent e) 
          {  System.out.println("  Connection disconnected"); }
          
          public void closed(ConnectionEvent e) 
          {  System.out.println("  Connection closed"); }
      });

      transport.addTransportListener( new TransportListener() {
          public void messageDelivered(TransportEvent e) 
          {  System.out.println("  Message delivered");  }
          
          public void messageNotDelivered(TransportEvent e) 
          {  System.out.println("  Message not delivered");  }
          
          public void messagePartiallyDelivered(TransportEvent e) 
          {  System.out.println("  Message partially delivered");  }
      });

      transport.connect(mailhost, port, user, password);
      transport.sendMessage(msg, msg.getAllRecipients());
      transport.close();
    }
    catch (Exception e) 
    {  reportSMTPExceptions(e);
       // System.out.println(e);
    }
  }  // end of sendEmail()



  private static void reportSMTPExceptions(Exception e)
  {
    if (e instanceof SendFailedException) {
      MessagingException sfe = (MessagingException) e;
      if (sfe instanceof SMTPSendFailedException) {
        SMTPSendFailedException ssfe = (SMTPSendFailedException) sfe;

        System.out.println("SMTP send failed:");
        System.out.println("  Command: " + ssfe.getCommand());
        System.out.println("  RetCode: " + ssfe.getReturnCode());
        System.out.println("  Response: " + ssfe.getMessage());
      }

      Exception ne;
      while (((ne = sfe.getNextException()) != null) && 
                                  (ne instanceof MessagingException)) {
        sfe = (MessagingException) ne;
        if (sfe instanceof SMTPAddressFailedException) {
          SMTPAddressFailedException ssfe = (SMTPAddressFailedException) sfe;
          System.out.println("Address failed:");
          System.out.println("  Address: " + ssfe.getAddress());
          System.out.println("  Command: " + ssfe.getCommand());
          System.out.println("  RetCode: " + ssfe.getReturnCode());
          System.out.println("  Response: " + ssfe.getMessage());
        }
        else if (sfe instanceof SMTPAddressSucceededException) {
          System.out.println("Address Succeeded:");
          SMTPAddressSucceededException ssfe = (SMTPAddressSucceededException) sfe;
          System.out.println("  Address: " + ssfe.getAddress());
          System.out.println("  Command: " + ssfe.getCommand());
          System.out.println("  RetCode: " + ssfe.getReturnCode());
          System.out.println("  Response: " + ssfe.getMessage());
        }
      }
    }
    else
      System.out.println(e);
  }  // end of reportSMTPExceptions()



  // ----------------- Java's Desktop API -------------------------


  public static void sendEmailByClient(String to, String subject, String body)
  {  sendEmailByClient(to, subject, body, null);  }


  public static void sendEmailByClient(String to, String subject,  
                                       String body, String fnm)
  /* Use the Desktop class to send e-mail using
     the OSes default e-mail client. javax.mail is NOT used.

     My e-mail client (Thunderbird) will not accept an attachment, 
     probably because "attachment" isn't a standard part of "mailto:".
     That's why I have a simpler version of sendEmailByClient() above.

     The mailto: spec. (RFC 2368):
        http://www.ietf.org/rfc/rfc2368.txt

     The client opens with its e-mail fields completed;
     the user must press send.
  */
  { 
    if (!Desktop.isDesktopSupported()) {
      System.out.println("Desktop mail not supported");
      return;
    }

    // construct "mailto:" string for Desktop.mail()
    String uriStr = String.format("mailto:%s?subject=%s&body=%s",
                              encodeMailto(to), encodeMailto(subject), 
                              encodeMailto(body) );
    if (fnm != null)
      uriStr += "&attachment=\"" + FileIO.getAbsolutePath(fnm) + "\"";   
           // this parameter isn't accepted by Thunderbird in the mail() call;
           // added quotes may work in Outlook, but I've not tested it
    try {
      Desktop desktop = Desktop.getDesktop();
      desktop.mail(new URI(uriStr));
    } 
    catch (Exception e) 
    {  System.out.println(e); }
  }  // end of sendEmailByClient()
  


  public static String encodeMailto(String str) 
  // from http://www.2ality.com/2010/12/simple-way-of-sending-emails-in-java.html
  {
    try {
      return URLEncoder.encode(str, "UTF-8").replace("+", "%20");
             // use URLEncoder and fix up "+" encoding
    } 
    catch (UnsupportedEncodingException e) {
       System.out.println("Could not encode: \"" + str + "\"");
       return null;
    }
  }  // end of encodeMailto()



  // ----------------- Java Process and batch script --------------------


  public static void sendEmailByTB(String to, String subject, String body, String fnm) 
  /* 
     Uses a batch script (TBExec.bat) to call Thunderbird through the command
     line; no use made of javax.mail or the Desktop API and "mailto:".

     Attachments are accepted unlike in sendEmailByClient().

     Thunderbird opens with its e-mail fields completed;
     the user must press send.
  */
  { String mailExec = String.format("cmd /c TBExec.bat %s \"%s\" \"%s\"",
                                                        to, subject, body);
    if (fnm != null)
      mailExec += " " + fnm;
          // no need for URI version of fnm after Thunderbird 2.0 
          // see http://kb.mozillazine.org/Command_line_arguments_(Thunderbird)
       
    // System.out.println(mailExec);
    try {
      Process p = Runtime.getRuntime().exec(mailExec);
      p.waitFor();
      System.out.println("Sent e-mail using Thunderbird");
    }
    catch (java.lang.Exception e) {
      System.out.println("Unable to send Thunderbird mail: " + e);
    }
  }  // end of sendEmailByTB()


}  // end of JMail class
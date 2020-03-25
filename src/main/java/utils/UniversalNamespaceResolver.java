
// UniversalNamespaceResolver.java
// http://www.ibm.com/developerworks/library/x-nmspccontext/

package utils;

import java.util.Iterator;

import javax.xml.XMLConstants;
import javax.xml.namespace.NamespaceContext;
import org.w3c.dom.Document;



public class UniversalNamespaceResolver implements NamespaceContext
{
  private Document sourceDoc;      // the delegate



  // stores the source document to search the namespaces in it
  public UniversalNamespaceResolver(Document document)
  {  sourceDoc = document;  }


  // The lookup for the namespace uris is delegated to the stored document
  public String getNamespaceURI(String prefix)
  {
    if (prefix.equals(XMLConstants.DEFAULT_NS_PREFIX))
      return sourceDoc.lookupNamespaceURI(null);
    else
      return sourceDoc.lookupNamespaceURI(prefix);
  }



  /* This method is not needed in this context, but can be implemented in a
     similar way.  */
  public String getPrefix(String namespaceURI)
  {  return sourceDoc.lookupPrefix(namespaceURI);  }


  public Iterator getPrefixes(String namespaceURI)
  { // not implemented yet
    return null;
  }

}  // end of UniversalNamespaceResolver class


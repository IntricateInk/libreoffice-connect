
// Registry.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2016


/* Based on RegistryTools.java in
   freedesktop/libreoffice/qadevOOo/runner/util/
   from http://code.metager.de/source/xref/
                     freedesktop/libreoffice/qadevOOo/runner/util/
*/

package utils;

import com.sun.star.uno.*;
import com.sun.star.registry.*;




public class Registry
{

  public static XSimpleRegistry open(String fnm,
               boolean forReadOnly, boolean canCreate)
  /* Opens registry file for reading/writing. If file doesn't
     exist a new one created.
  */
  {
    XSimpleRegistry reg = Lo.createInstanceMCF(XSimpleRegistry.class, 
                              "com.sun.star.registry.SimpleRegistry");
    if (reg == null)
      return null;
    try {
      String urlStr = FileIO.fnmToURL(fnm);
      reg.open(urlStr, forReadOnly, canCreate);
      System.out.println("Openned registry file: " + fnm);
      return reg;
    }
    catch(com.sun.star.uno.Exception e) {
      System.out.println("Could not open registry file: " + fnm);
      System.out.println(e);
      return null;
    }
  }  // end of open()



  public static XSimpleRegistry view(String fnm)
  {  return open(fnm, true, false);  }



  public static void close(XSimpleRegistry reg)
  {
    if (reg == null)
      return;
    try {
      reg.close();
    }
    catch (InvalidRegistryException e) {
      System.out.print("Couldn't close registry: " + e);
    }
  }  // end of close()



  public static boolean compareKeys(XRegistryKey key1, XRegistryKey key2)
  /* Compares two registry keys, their names, value types and values.
     return true if key names, value types
     and values are equal, else returns false.
  */
  {
    if (key1 == null || key2 == null || !key1.isValid() || !key2.isValid())
      return false;

    String keyName1 = getShortKeyName(key1.getKeyName());
    String keyName2 = getShortKeyName(key2.getKeyName());
    if (!keyName1.equals(keyName2)) 
      return false;

    try {
      if (key1.getValueType() != key2.getValueType()) return false;
    }
    catch (InvalidRegistryException e) {
      return false;
    }

    RegistryValueType type;
    try {
      type = key1.getValueType();
      if (type.equals(RegistryValueType.ASCII)) {
        if (!key1.getAsciiValue().equals(key2.getAsciiValue()))
          return false;
      }
      else if (type.equals(RegistryValueType.STRING)) {
        if (!key1.getStringValue().equals(key2.getStringValue()))
          return false;
      }
      else if (type.equals(RegistryValueType.LONG)) {
        if (key1.getLongValue() != key2.getLongValue())
          return false;
      }
      else if (type.equals(RegistryValueType.BINARY)) {
        byte[] bin1 = key1.getBinaryValue();
        byte[] bin2 = key2.getBinaryValue();
        if (bin1.length != bin2.length)
          return false;
        for (int i = 0; i < bin1.length; i++)
          if (bin1[i] != bin2[i]) return false;
      }
      else if (type.equals(RegistryValueType.ASCIILIST)) {
        String[] list1 = key1.getAsciiListValue();
        String[] list2 = key2.getAsciiListValue();
        if (list1.length != list2.length)
          return false;
        for (int i = 0; i < list1.length; i++)
          if (!list1[i].equals(list2[i])) return false;
      }
      else if (type.equals(RegistryValueType.STRINGLIST)) {
        String[] list1 = key1.getStringListValue();
        String[] list2 = key2.getStringListValue();
        if (list1.length != list2.length)
          return false;
        for (int i = 0; i < list1.length; i++)
          if (!list1[i].equals(list2[i])) return false;
      }
      else if (type.equals(RegistryValueType.LONGLIST)) {
        int[] list1 = key1.getLongListValue();
        int[] list2 = key2.getLongListValue();
        if (list1.length != list2.length)
          return false;
        for (int i = 0; i < list1.length; i++)
          if (list1[i] != list2[i]) return false;
      }
    }
    catch (com.sun.star.uno.Exception e) {
      return false;
    }
    return true;
  }  // end of compareKeys()



  public static String getShortKeyName(String keyName)
  /* Gets name of the key relative to its parent.
     For example if full name of key is '/key1/subkey'
     short key name is 'subkey'
  */
  {
    if (keyName == null) 
      return null;
    int idx = keyName.lastIndexOf("/");
    if (idx < 0) 
      return keyName;
    else 
      return keyName.substring(idx + 1);
  }  // end of getShortKeyName()



  public static boolean compareKeyTrees(XRegistryKey tree1, 
                                XRegistryKey tree2, boolean compareRoot)
  /* Compare all child keys.
     return true if keys and their sub keys are equal.
   */
  {
    if (compareRoot && !compareKeys(tree1, tree2)) 
      return false;
    try {
      String[] keyNames1 = tree1.getKeyNames();
      String[] keyNames2 = tree2.getKeyNames();
      if (keyNames1 == null && keyNames2 == null) 
        return true;

      if (keyNames1 == null || keyNames2 == null || 
          keyNames2.length != keyNames1.length)
        return false;

      for (int i = 0; i < keyNames1.length; i++) {
        String keyName = getShortKeyName(keyNames1[i]);
        XRegistryKey key2 = tree2.openKey(keyName);
        if (key2 == null)
          // key with the same name doesn't exist in the second tree
          return false;

        if (!tree1.getKeyType(keyName).equals(tree2.getKeyType(keyName)))
          return false;

        if (tree1.getKeyType(keyName).equals(RegistryKeyType.LINK)) {
          if (!getShortKeyName(tree1.getLinkTarget(keyName)).equals(
              getShortKeyName(tree2.getLinkTarget(keyName))))
            return false;
        }
        else {
          if (!compareKeyTrees(tree1.openKey(keyName), 
                               tree2.openKey(keyName), true)) 
            return false;
        }
      }
    }
    catch (InvalidRegistryException e) {
      return false;
    }
    return true;
  }  // end of compareKeyTrees()



  public static boolean compareKeyTrees(XRegistryKey tree1, XRegistryKey tree2)
  // return true if keys and their sub keys are equal
  {  return compareKeyTrees(tree1, tree2, false);  }



  public static void print(XSimpleRegistry reg)
  /* Prints all keys and subkeys information
     (key name, type, value, link target, attributes)  */
  {
    try {
      print(reg.getRootKey());
    }
    catch (InvalidRegistryException e) {
      System.out.println("Cannot open root registry key");
    }
  }  // end of print()



  public static void print(XRegistryKey root)
  {
    if (root == null) {
      System.out.println("/(null)");
      return;
    }
    System.out.println("/");
    try {
      printRegTree(root, "  ");
    }
    catch (InvalidRegistryException e) {
      System.out.println("Could not access registry: " + e);
    }
  }  // end of print()



  private static void printRegTree(XRegistryKey key, String margin)
                                throws InvalidRegistryException
  {
    String[] subKeys = key.getKeyNames();
    if (subKeys == null || subKeys.length == 0) 
      return;
    for (int i = 0; i < subKeys.length; i++) {
      printRegKey(key, subKeys[i], margin);
      XRegistryKey subKey = key.openKey(getShortKeyName(subKeys[i]));
      printRegTree(subKey, margin + "  ");
      subKey.closeKey();
    }
  }  // end of printRegTree()



  private static void printRegKey(XRegistryKey parentKey,
                               String keyName, String margin)
                                     throws InvalidRegistryException
  { System.out.print(margin);
    keyName = getShortKeyName(keyName);
    XRegistryKey key = parentKey.openKey(keyName);
    if (key != null)
      System.out.print("/" + getShortKeyName(key.getKeyName()) + " ");
    else {
      System.out.println("(null)");
      return;
    }

    if (!key.isValid()) {
      System.out.println("(not valid)");
      return;
    }

    if (key.isReadOnly())
      System.out.print("(read only) ");
    if (parentKey.getKeyType(keyName) == RegistryKeyType.LINK) {
      System.out.println("(link to " + parentKey.getLinkTarget(keyName) + ")");
      return;
    }

    RegistryValueType type;
    try {
      type = key.getValueType();
      if (type.equals(RegistryValueType.ASCII))
        System.out.println("[ASCII] = '" + key.getAsciiValue() + "'");
      else if (type.equals(RegistryValueType.STRING))
        System.out.println("[STRING] = '" + key.getStringValue() + "'");
      else if (type.equals(RegistryValueType.LONG))
        System.out.println("[LONG] = " + key.getLongValue());
      else if (type.equals(RegistryValueType.BINARY)) {
        byte[] bin = key.getBinaryValue();
        System.out.println("[BINARY], size = " + bin.length);
        //for (int i = 0; i < bin.length; i++)
        //  System.out.print("" + bin[i] + ",");
      }
      else if (type.equals(RegistryValueType.ASCIILIST)) {
        System.out.print("[ASCIILIST] = {");
        String[] list = key.getAsciiListValue();
        for (int i = 0; i < list.length; i++)
          System.out.print("'" + list[i] + "',");
        System.out.println("}");
      }
      else if (type.equals(RegistryValueType.STRINGLIST)) {
        System.out.print("[STRINGLIST] = {");
        String[] list = key.getStringListValue();
        for (int i = 0; i < list.length; i++)
          System.out.print("'" + list[i] + "',");
        System.out.println("}");
      }
      else if (type.equals(RegistryValueType.LONGLIST)) {
        System.out.print("[LONGLIST] = {");
        int[] list = key.getLongListValue();
        for (int i = 0; i < list.length; i++)
          System.out.print("" + list[i] + ",");
        System.out.println("}");
      }
      else
        System.out.println("");
    }
    catch (com.sun.star.uno.Exception e) {
      System.out.println("Cannot print registry key: " + e);
    } 
    finally {
      key.closeKey();
    }
  }  // end of printRegKey()


}  // end of Registry class

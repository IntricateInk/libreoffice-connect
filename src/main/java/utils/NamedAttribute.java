
// NamedAttribute.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

// Used by JPrint.java

package utils;

import javax.print.attribute.Attribute;


public class NamedAttribute implements Comparable<NamedAttribute>
{
  private String name;
  private Attribute attr;


  public NamedAttribute(Attribute attr)
  {
    name = attr.getName();
    this.attr = attr;
  }

  public String getName()
  {  return name;  }

  public Attribute getAttribute()
  {  return attr;  }


  public boolean equals(Object obj)
  {
    if(this == obj)
      return true;
    if((obj == null) || (obj.getClass() != this.getClass()))
      return false;

    NamedAttribute na = (NamedAttribute)obj;
    return name.equals(na.getName());
  }


  public int hashCode() 
  {  return name.hashCode();  }


  public int compareTo(NamedAttribute na) 
  // ascending order by name
  {  return name.compareTo(na.getName());  }


  public String toString() 
  {  return name; }


}  // end of NamedAttribute class

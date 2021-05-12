
package javaapplication2;


public class rollNotFoundException extends Exception {
    
    String msg;
    rollNotFoundException(String m)
    {
        msg=m;
    }
    
    public String toString()
    {
        return msg;
    }
 }

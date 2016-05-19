package xtr.util;

public class StringUtil {  
  
    public static String strLeft(final String text, String subtext) {  
        if(!hasText(text) || !hasText(subtext)) {  
            return "";  
        }  
          
        int find = text.indexOf(subtext);  
        return (find!=-1) ? text.substring(0, find) : "";  
    }  
      
    public static String strRight(final String text, String subtext) {  
        if(!hasText(text) || !hasText(subtext)) {  
            return "";  
        }  
          
        int find = text.indexOf(subtext); 
        int lg = subtext.length() ;
        int qr = find + lg;
        return (find!=-1) ? text.substring(qr) : "";  
    }  
  
    public static String strLeftBack(final String text, String subtext) {  
        if(!hasText(text) || !hasText(subtext)) {  
            return "";  
        }  
          
        int find = text.lastIndexOf(subtext);  
        return (find!=-1) ? text.substring(0, find) : "";  
    }  
      
    public static String strRightBack(final String text, String subtext) {  
        if(!hasText(text) || !hasText(subtext)) {  
            return "";  
        }  
          
        int find = text.lastIndexOf(subtext);  
        return (find!=-1) ? text.substring(find+1) : "";  
    }  
      
    private static boolean hasText(String text) {  
        return (text!=null) && (!"".equals(text));  
    }  
}  
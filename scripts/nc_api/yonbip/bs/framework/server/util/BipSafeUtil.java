package yonbip.bs.framework.server.util;
public class BipSafeUtil implements java.io.Serializable {
  private static final long serialVersionUID = 1L;
  private static final BipSafeUtil INSTANCE = new BipSafeUtil();
  public static BipSafeUtil getInstance() { return INSTANCE; }
  public boolean checkBackList(String s) { return false; }
}


import java.io.*;
import java.util.*;
import nc.bs.framework.common.*;
import nc.login.bs.INCLoginService;
import nc.login.vo.*;
import nc.identityverify.vo.AuthenSubject;
import nc.itf.uap.IUAPQueryBS;
import nc.jdbc.framework.processor.MapListProcessor;

public class NcSql {
  static String esc(Object v) {
    if (v == null) return "";
    String s = String.valueOf(v);
    if (s.indexOf(',') >= 0 || s.indexOf('"') >= 0 || s.indexOf('\n') >= 0) return '"' + s.replace("\"", "\"\"") + '"';
    return s;
  }
  public static void main(String[] a) throws Exception {
    String url = a[0], user = a[1], pwd = a[2], bc = a[3];
    String sqlFile = a[4], outFile = a[5];
    String sql = new String(java.nio.file.Files.readAllBytes(new File(sqlFile).toPath()), "UTF-8").replace("\uFEFF", "").trim();
    if (sql.endsWith(";")) sql = sql.substring(0, sql.length()-1);
    RuntimeEnv.getInstance().setProperty("SERVICEDISPATCH_URL", url);
    NCLocator loc = NCLocator.getInstance();
    UserExit ue = UserExit.getInstance();
    ue.setKeepToken(true); ue.setUserCode(user); ue.setUserDataSource("design"); ue.setBizCenterCode(bc);
    InvocationInfoProxy ip = InvocationInfoProxy.getInstance();
    ip.reset(); ip.setBizDateTime(System.currentTimeMillis());
    INCLoginService ls = (INCLoginService) loc.lookup(INCLoginService.class);
    LoginRequest req = new LoginRequest();
    req.setUserCode(user); req.setUserPWD(pwd); req.setBusiCenterCode(bc); req.setLangCode("simpchn"); req.setDeviceType(0);
    AuthenSubject subj = new AuthenSubject();
    subj.setUserCode(user); subj.setUserPWD(req.getUserPWD()); subj.setBusiCenterCode(bc);
    req.putAttachProp(AuthenSubject.class.getName(), subj);
    LoginResponse resp = ls.login(req, false);
    NCSession s = resp.getNcSession();
    if (s == null || s.getSessionID() == null) { System.err.println("LOGIN_FAILED result=" + resp.getLoginResult()); System.exit(2); }
    ue.setUserCode(s.getUserCode()); ue.setUserDataSource(s.getDsName()); ue.setBizCenterCode(s.getBusiCenterCode());
    ip.setUserId(s.getUserID()); ip.setUserCode(s.getUserCode()); ip.setUserDataSource(s.getDsName()); ip.setBizCenterCode(s.getBusiCenterCode()); ip.setLangCode(s.getLangCode());
    IUAPQueryBS q = (IUAPQueryBS) loc.lookup(IUAPQueryBS.class);
    Object r = q.executeQuery(sql, null, new MapListProcessor());
    List rows = (r instanceof List) ? (List) r : new ArrayList();
    PrintWriter w = new PrintWriter(new OutputStreamWriter(new FileOutputStream(outFile), "UTF-8"));
    if (!rows.isEmpty()) {
      Map m0 = (Map) rows.get(0);
      List<String> keys = new ArrayList<String>();
      for (Object k : m0.keySet()) keys.add(String.valueOf(k));
      Collections.sort(keys);
      for (int i=0;i<keys.size();i++) { if (i>0) w.print(','); w.print(esc(keys.get(i))); }
      w.println();
      for (Object o : rows) {
        Map m = (Map) o;
        for (int i=0;i<keys.size();i++) { if (i>0) w.print(','); w.print(esc(m.get(keys.get(i)))); }
        w.println();
      }
    }
    w.close();
    System.out.println("ROWS=" + rows.size() + " OUT=" + outFile);
  }
}

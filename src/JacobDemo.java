import java.io.File;
import java.util.HashMap;
import java.util.Map;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
 
/*
 * 实现文档格式之间的转换，我使用的是jacob-1.7版本，需要jacob.jar来调用activex控件;
 * 本机需安装WPS/office，还需要jacob.dll 
 * jacob.dll 需要放置在系统system32下，如果系统是c盘：C://windows/system32/下面;
 * jacob.dll放在类似这样的目录下，D:\Java\jre1.8.0_31\bin;
 * 
 */
public class JacobDemo {
    
    // word运行程序对象
    private ActiveXComponent app = null;
    // 所有文档集合(文档容器)
    private Dispatch documents = null;
    // 当前要打开的文档
    private Dispatch doc = null;
    // 开始查找位置
    private static Dispatch selection = null;
    // 设置是否保存后才退出的标志
    private boolean saveOnExit = false;
    
    // 转换格式 
    private static final int wdFormatPDF = 17;
    private static final int xlTypePDF = 0;  
    private static final int ppSaveAsPDF = 32;
 
    //word转pdf
    public boolean wordToPdf(String source, String target) { 
        try {  
            ComThread.InitSTA();
            // 打开word应用程序  
            app = new ActiveXComponent("Word.Application");
            // 设置word不可见
            app.setProperty("Visible", new Variant(false));
            // 获得word中所有打开的文档,返回Documents对象  
            documents = app.getProperty("Documents").toDispatch();
            //打开文档  
            doc = Dispatch.call(documents, "Open", source, false, true).toDispatch();
            //word转pdf
            delFile(new File(target));
            Dispatch.call(doc, "SaveAs", target, wdFormatPDF);  
            System.out.println("word转换pdf文档："+ target);
            
            return true;  
        } catch (Exception e) {  
            System.out.println("Word转PDF出错：" + e.getMessage());  
            return false;  
        } finally {  
            close();
        }  
    }
    
    // excel转换为pdf  
    public boolean excelToPdf(String source, String target) {  
        try {  
            ComThread.InitSTA();
            // 打开excel应用程序
            app = new ActiveXComponent("Excel.Application");  
            // 设置excel不可见
            app.setProperty("Visible", false);  
            // 获得excel中所有打开的文档,返回documents对象  
            documents = app.getProperty("Workbooks").toDispatch(); 
            // 打开文档  
            doc = Dispatch.call(documents, "Open", source, false, true).toDispatch();  
            // excel转换为pdf 
            delFile(new File(target));
            Dispatch.call(doc, "ExportAsFixedFormat", target, xlTypePDF);  
            System.out.println("excel转换pdf文档："+ target);
            
            return true;  
        } catch (Exception e) {
            System.out.println("excel转PDF出错：" + e.getMessage());
            return false;  
        }finally{
            close();
        } 
    }
    
    // ppt转换为pdf  
    public boolean pptToPdf(String source, String target) {  
        try {  
            ComThread.InitSTA();
            // 打开应用程序  
            app = new ActiveXComponent("PowerPoint.Application");  
            // 获得ppt中所有打开的文档,返回Documents对象   
            documents = app.getProperty("Presentations").toDispatch();  
            // 打开文档  
            doc = Dispatch.call(documents, "Open", source, true,// ReadOnly  
                    true,// Untitled指定文件是否有标题  
                    false// WithWindow指定文件是否可见  
                    ).toDispatch();  
 
            // ppt转换为pdf
            delFile(new File(target));
            Dispatch.call(doc, "SaveAs", target, ppSaveAsPDF);  
            System.out.println("ppt转换pdf文档："+ target);
            
            return true;  
        } catch (Exception e) { 
            System.out.println("ppt转PDF出错：" + e.getMessage()); 
            return false;  
        }finally{
            close();
        }
    }
    
    //根据模板新建文档
    public boolean template(String template, String target, Map<String, String> map){
        try {  
            ComThread.InitSTA();
            // 打开应用程序  
            app = new ActiveXComponent("Word.Application");
            // 设置word不可见
            app.setProperty("Visible", new Variant(false));
            // 获得word中所有打开的文档,返回documents对象  
            documents = app.getProperty("Documents").toDispatch();
            //打开文档  
            doc = Dispatch.call(documents, "Open", template, false, true).toDispatch();
            //设置模板参数
            setTemplate(map);
            
            delFile(new File(target));
            //另存为pdf
            Dispatch.call(doc, "SaveAs", target, wdFormatPDF);
            //另存为word
            //saveAs(target);
            
            System.out.println("根据模板新建文档："+ target);
            
            return true;  
        } catch (Exception e) {  
            System.out.println("根据模板新建文档出错：" + e.getMessage());  
            return false;  
        } finally {  
            close();
        }
        
    }
    
    // 设置插入点为文件首位置
    public void moveStart() {
        if (selection == null){
            selection = Dispatch.get(app, "Selection").toDispatch();
        }
            
        Dispatch.call(selection, "HomeKey", new Variant(6));
    }
    
    /**
     * 从选定内容或插入点开始查找文本
     * @param toFindText 要查找的文本
     * @return boolean true-查找到并选中该文本，false-未查找到文本
     */
    public boolean find(String toFindText) {
       
        if (toFindText == null || toFindText.equals("")){
            return false;
        }
        
        // 从selection所在位置开始查询
        Dispatch find = Dispatch.call(selection, "Find").toDispatch();
        // 设置要查找的内容
        Dispatch.put(find, "Text", toFindText);
        // 向前查找
        Dispatch.put(find, "Forward", "True");
        // 设置格式
        Dispatch.put(find, "Format", "True");
        // 大小写匹配
        Dispatch.put(find, "MatchCase", "True");
        // 全字匹配
        Dispatch.put(find, "MatchWholeWord", "True");
        // 查找并选中
        return Dispatch.call(find, "Execute").getBoolean();
    }
    
    // 设置模板参数
    public void setTemplate(Map<String, String> map){
        for (String key : map.keySet()) {
            moveStart(); //设置查找位置
            String value = map.get(key);
            if (find(key)){
                //替换文本
                Dispatch.put(selection, "Text", value);
            }
        }
    }
    
    // 删除文件
    public void delFile(File file){
        if (file.exists()) {  
            file.delete();  
        }
    }
    
    // 保存文档
    public void save(String savePath) {
        Dispatch.call(doc, "SaveAs", savePath);
    }
    
    // 另保存文档
    public void saveAs(String savePath) {
        Dispatch.call(Dispatch.call(app, "WordBasic").getDispatch(), "FileSaveAs", savePath);
    }
    
    //关闭当前word文档
    public void closeDocument() {
        if (doc != null) {
            Dispatch.call(doc, "Close", saveOnExit);
            doc = null;
        }
    }
 
    //关闭全部应用
    public void close() {
        closeDocument();
        if (app != null) {
            Dispatch.call(app, "Quit");
            app = null;
        }
        selection = null;
        documents = null;
        ComThread.Release();
    }
 
    
    public static void main(String[] args) {
        
        JacobDemo jd = new JacobDemo();
        Map<String, String> map = new HashMap<String, String>();
        map.put("<biaoti>", "这里是测试的标题");
        map.put("<nizhi>", "王蒙");
        map.put("<shenhe>", "陈留飞");
        map.put("<huiqian>", "刘玉栋，\r\n刘育栋，\r\n刘思涵");
        map.put("<biaozhunhua>", "冥界");
        map.put("<pizhun>", "昊天");
        
        //jd.wordToPdf("E:/dianziqianzhang/数据储蓄协议.docx", "E:/dianziqianzhang/数据储蓄协议.pdf");
        jd.template("C:/Users/DELL/Desktop/test/templates/design_template.docx", "C:/Users/DELL/Desktop/test/templates/design_template.pdf", map);
    }
}

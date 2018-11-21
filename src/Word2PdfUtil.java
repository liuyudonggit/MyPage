import java.io.File;
import java.io.IOException;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class Word2PdfUtil
{
    
    static final int wdDoNotSaveChanges = 0;// 不保存待定的更改。
    
    private static final int wdFormatPDF = 17;// word转PDF 格式
    
    private static final int ppFormatPDF = 32;// ppt转pdf
    
    private static final int xlFormatPDF = 0;// excel转pdf
    
    public static void main(String[] args)
        throws IOException
    {
        String source1 = "C:\\Users\\DELL\\Desktop\\test\\design_template.docx";
        String xlsx = "C:\\Users\\DELL\\Desktop\\test\\Book1.xlsx";
        String ppt = "C:\\Users\\DELL\\Desktop\\test\\ppt.pptx";
        String target1 = "C:\\Users\\DELL\\Desktop\\testpdf\\design_template.pdf";
        String xlsxPDF = "C:\\Users\\DELL\\Desktop\\testpdf\\excel.pdf";
        String pptPDF = "C:\\Users\\DELL\\Desktop\\testpdf\\ppt.pdf";
        Word2PdfUtil pdf = new Word2PdfUtil();
        pdf.word2pdf(source1, target1);
        pdf.excel2PDF(xlsx, xlsxPDF);
        pdf.ppt2PDF(ppt, pptPDF);
    }
    
    public boolean word2pdf(String source, String target)
    {
        System.out.println("Word转PDF开始启动...");
        long start = System.currentTimeMillis();
        ActiveXComponent app = null;
        try
        {
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", false);
            Dispatch docs = app.getProperty("Documents").toDispatch();
            System.out.println("打开文档：" + source);
            Dispatch doc = Dispatch.call(docs, "Open", source, false, true).toDispatch();
            System.out.println("转换文档到PDF：" + target);
            File tofile = new File(target);
            if (tofile.exists())
            {
                tofile.delete();
            }
            Dispatch.call(doc, "SaveAs", target, wdFormatPDF);// 转word
            Dispatch.call(doc, "Close", false);
            long end = System.currentTimeMillis();
            System.out.println("转换完成，用时：" + (end - start) + "ms");
            return true;
        }
        catch (Exception e)
        {
            System.out.println("Word转PDF出错：" + e.getMessage());
            return false;
        }
        finally
        {
            if (app != null)
            {
                app.invoke("Quit", wdDoNotSaveChanges);
            }
        }
    }
    
    /**
         * ppt文档转换
         * 
         * @param inputFile
         * @param pdfFile
         * @author SHANHY
         */
    private boolean ppt2PDF(String inputFile, String pdfFile)
    {
        ComThread.InitSTA();
        
        long start = System.currentTimeMillis();
        ActiveXComponent app = null;
        Dispatch ppt = null;
        try
        {
            app = new ActiveXComponent("PowerPoint.Application");// 创建一个PPT对象
            // app.setProperty("Visible", new Variant(false)); // 不可见打开（PPT转换不运行隐藏，所以这里要注释掉）
            // app.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏
            Dispatch ppts = app.getProperty("Presentations").toDispatch();// 获取文挡属性
            
            System.out.println("打开文档 >>> " + inputFile);
            // 调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
            ppt = Dispatch.call(ppts, "Open", inputFile, true,// ReadOnly
                true,// Untitled指定文件是否有标题
                false// WithWindow指定文件是否可见
            )
                .toDispatch();
            
            System.out.println("转换文档 [" + inputFile + "] >>> [" + pdfFile + "]");
            Dispatch.call(ppt, "SaveAs", pdfFile, ppFormatPDF);
            Dispatch.call(ppt, "Close", false);
            long end = System.currentTimeMillis();
            
            System.out.println("用时：" + (end - start) + "ms.");
            
            return true;
            
        }
        catch (Exception e)
        {
            e.printStackTrace();
            System.out.println("========Error:文档转换失败：" + e.getMessage());
        }
        finally
        {
            System.out.println("关闭文档");
            if (app != null)
                app.invoke("Quit", new Variant[] {});
        }
        
        ComThread.Release();
        ComThread.quitMainSTA();
        return false;
    }
    
    /**
         * Excel文档转换
         * 
         * @param inputFile
         * @param pdfFile
         * @author SHANHY
         */
    private boolean excel2PDF(String inputFile, String pdfFile)
    {
        ComThread.InitSTA();
        
        long start = System.currentTimeMillis();
        ActiveXComponent app = null;
        Dispatch excel = null;
        try
        {
            app = new ActiveXComponent("Excel.Application");// 创建一个PPT对象
            app.setProperty("Visible", new Variant(false));
            app.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏
            Dispatch excels = app.getProperty("Workbooks").toDispatch();// 获取文挡属性
            
            System.out.println("打开文档 >>> " + inputFile);
            // 调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
            excel = Dispatch.call(excels, "Open", inputFile, false, true).toDispatch();
            // 调用Document对象方法，将文档保存为pdf格式
            System.out.println("转换文档 [" + inputFile + "] >>> [" + pdfFile + "]");
            // Excel 不能调用SaveAs方法
            Dispatch.call(excel, "ExportAsFixedFormat", xlFormatPDF, pdfFile);
            //Dispatch.call(excel, "SaveAs", pdfFile, wdFormatPDF);// 转word
            Dispatch.call(excel, "Close", false);
            long end = System.currentTimeMillis();
            
            System.out.println("用时：" + (end - start) + "ms.");
            return true;
        }
        catch (Exception e)
        {
            e.printStackTrace();
            System.out.println("========Error:文档转换失败：" + e.getMessage());
        }
        finally
        {
            if (app != null)
                app.invoke("Quit", new Variant[] {});
        }
        
        ComThread.Release();
        ComThread.quitMainSTA();
        return false;
    }
    
}

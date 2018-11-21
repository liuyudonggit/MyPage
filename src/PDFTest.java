import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.net.URLDecoder;
import java.text.SimpleDateFormat;
import java.util.Properties;

import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;

import demo.CommandUtil;
import demo.FileConstants;
import demo.FileUtil;


public class PDFTest
{
    
    public static void main(String[] args)
    {
        PDFTest ptest = new PDFTest();
        try
        {
           //ptest.drawingToPDF("A0.dwg");//测试DWG转PDF
          
           String currentTime = getCurrentTime();
           ptest.generatePDF("E:\\DrawingToPDF\\AutoCAD\\Upload\\A4X4.pdf","E:\\DrawingToPDF\\AutoCAD\\generatePDF\\A4X4"+currentTime+".pdf");
        }
        catch (Exception e)
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        
    }
    

    private static String getCurrentTime()
    {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddhhmmss");
        return sdf.format(System.currentTimeMillis()); 
    }


    /**  
    * @Title: generatePDF 
    * @Description: pdf签名 
    * @param pdfPath 原pdf地址
    * @param newPDFPath 新pdf存储地址
    * @param format 类型 A0,A1...
    * @throws Exception       
    * @return void    返回类型  
    * @throws  
    * @date 2018-11-13
    * @author liuyudong@glaway.com
    */
    public void generatePDF(String pdfPath, String newPDFPath) throws Exception
    {
        PdfReader reader = null;
        PdfStamper stamper = null;
        try
        {
             reader = new PdfReader(pdfPath);
           // ByteArrayOutputStream bos = new ByteArrayOutputStream();
             stamper = new PdfStamper(reader,  new FileOutputStream(newPDFPath));
            int totalPageNumber = reader.getNumberOfPages();//获取页数
           // AcroFields form = stamper.getAcroFields();
            // pageNumber = 1
            //String textFromPage = PdfTextExtractor.getTextFromPage(reader, 1);
            //Document document = new Document(); 
            //PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(newPDFPath));  
            //for(int pageNumber=1;pageNumber<=totalPageNumber;pageNumber++){
                
                 PdfContentByte cb = stamper.getOverContent(1);  
                 BaseFont baseFont = BaseFont.createFont("STSongStd-Light","UniGB-UCS2-H",BaseFont.NOT_EMBEDDED);  
                 cb.beginText();  
                //设置坐标
                 String format = FileUtil.getInfo(pdfPath);
                 
                 
                 switch (format)
                {
                    case "A0":
                        cb.setFontAndSize(baseFont, 10);  
                        cb.setTextMatrix(2890, 148);  //Y轴间隔14单位
                        cb.showText("拟制  720, 132");  
                        cb.setTextMatrix(2890, 134);  
                        cb.showText("审核  720, 118");  
                        cb.setTextMatrix(2890, 120); 
                        cb.showText("工艺1  720, 104");  
                        cb.setTextMatrix(2890, 106); 
                        cb.showText("工艺2  720, 90");  
                        cb.setTextMatrix(2890, 92); 
                        cb.showText("工艺3  720, 76");  
                        cb.setTextMatrix(2890, 78); 
                        cb.showText("工艺4  720, 62");  
                        cb.setTextMatrix(2890, 64); 
                        cb.showText("工艺5  720, 48"); 
                        cb.setTextMatrix(2890, 50);  
                        cb.showText("标准化   2890, 50"); 
                        cb.setTextMatrix(2890, 36);  
                        cb.showText("批准   2890, 36"); 
                        break;
                    case "A1":
                        cb.setFontAndSize(baseFont, 10);  
                        cb.setTextMatrix(1900, 148);  //Y轴间隔14单位
                        cb.showText("拟制  720, 132");  
                        cb.setTextMatrix(1900, 134);  
                        cb.showText("审核  720, 118");  
                        cb.setTextMatrix(1900, 120); 
                        cb.showText("工艺1  720, 104");  
                        cb.setTextMatrix(1900, 106); 
                        cb.showText("工艺2  720, 90");  
                        cb.setTextMatrix(1900, 92); 
                        cb.showText("工艺3  720, 76");  
                        cb.setTextMatrix(1900, 78); 
                        cb.showText("工艺4  720, 62");  
                        cb.setTextMatrix(1900, 64); 
                        cb.showText("工艺5  720, 48"); 
                        cb.setTextMatrix(1900, 50);  
                        cb.showText("标准化   2890, 50"); 
                        cb.setTextMatrix(1900, 36);  
                        cb.showText("批准   2890, 36"); 
                        break;
                    case "A2":
                        cb.setFontAndSize(baseFont, 10);  
                        cb.setTextMatrix(1200, 146);  //Y轴间隔14单位
                        cb.showText("拟制  720, 132");  
                        cb.setTextMatrix(1200, 132);  
                        cb.showText("审核  720, 118");  
                        cb.setTextMatrix(1200, 118); 
                        cb.showText("工艺1  720, 104");  
                        cb.setTextMatrix(1200, 104); 
                        cb.showText("工艺2  720, 90");  
                        cb.setTextMatrix(1200, 90); 
                        cb.showText("工艺3  720, 76");  
                        cb.setTextMatrix(1200, 76); 
                        cb.showText("工艺4  720, 62");  
                        cb.setTextMatrix(1200, 62); 
                        cb.showText("工艺5  720, 48"); 
                        cb.setTextMatrix(1200, 48);  
                        cb.showText("标准化   2890, 50"); 
                        cb.setTextMatrix(1200, 34);  
                        cb.showText("批准   2890, 36"); 
                        break;
                    case "A3":
                        cb.setFontAndSize(baseFont, 10);  
                        cb.setTextMatrix(720, 132);  //Y轴间隔14单位
                        cb.showText("拟制  720, 132");  
                        cb.setTextMatrix(720, 118);  
                        cb.showText("审核  720, 118");  
                        cb.setTextMatrix(720, 104); 
                        cb.showText("工艺1  720, 104");  
                        cb.setTextMatrix(720, 90); 
                        cb.showText("工艺2  720, 90");  
                        cb.setTextMatrix(720, 76); 
                        cb.showText("工艺3  720, 76");  
                        cb.setTextMatrix(720, 62); 
                        cb.showText("工艺4  720, 62");  
                        cb.setTextMatrix(720, 48); 
                        cb.showText("工艺5  720, 48"); 
                        cb.setTextMatrix(720, 34);  
                        cb.showText("标准化   720, 34"); 
                        cb.setTextMatrix(720, 20);  
                        cb.showText("批准   720, 20"); 
                        break;
                    case "A4":
                        cb.setFontAndSize(baseFont, 10);  
                        cb.setTextMatrix(130, 132);  //Y轴间隔14单位
                        cb.showText("拟制  130, 132");  
                        cb.setTextMatrix(130, 118);  
                        cb.showText("审核  130, 118");  
                        cb.setTextMatrix(130, 104); 
                        cb.showText("工艺1  130, 104");  
                        cb.setTextMatrix(130, 90); 
                        cb.showText("工艺2  130, 90");  
                        cb.setTextMatrix(130, 76); 
                        cb.showText("工艺3  130, 76");  
                        cb.setTextMatrix(130, 62); 
                        cb.showText("工艺4  130, 62");  
                        cb.setTextMatrix(130, 48); 
                        cb.showText("工艺5  130, 48"); 
                        cb.setTextMatrix(130, 34);  
                        cb.showText("标准化   130, 34"); 
                        cb.setTextMatrix(130, 20);  
                        cb.showText("批准   130, 20"); 
                        cb.setTextMatrix(420, 76); 
                        cb.showText("版本  420, 76");  
                        break;
                    case "A1×3":
                        cb.setFontAndSize(baseFont, 10);  
                        cb.setTextMatrix(720, 132);  //Y轴间隔14单位
                        cb.showText("拟制  720, 132");  
                        cb.setTextMatrix(720, 118);  
                        cb.showText("审核  720, 118");  
                        cb.setTextMatrix(720, 104); 
                        cb.showText("工艺1  720, 104");  
                        cb.setTextMatrix(720, 90); 
                        cb.showText("工艺2  720, 90");  
                        cb.setTextMatrix(720, 76); 
                        cb.showText("工艺3  720, 76");  
                        cb.setTextMatrix(720, 62); 
                        cb.showText("工艺4  720, 62");  
                        cb.setTextMatrix(720, 48); 
                        cb.showText("工艺5  720, 48"); 
                        cb.setTextMatrix(720, 34);  
                        cb.showText("标准化   720, 34"); 
                        cb.setTextMatrix(720, 20);  
                        cb.showText("批准   720, 20"); 
                        break;
                    case "A1×4":
                        cb.setFontAndSize(baseFont, 1);  
                        cb.setTextMatrix(780, 18.5f);  //Y轴间隔1.75单位
                        cb.showText("拟制  720, 132");  
                        cb.setTextMatrix(780, 16.75f);  
                        cb.showText("审核  720, 118");  
                        cb.setTextMatrix(780, 15); 
                        cb.showText("工艺1  720, 104");  
                        cb.setTextMatrix(780, 13.25f); 
                        cb.showText("工艺2  720, 90");  
                        cb.setTextMatrix(780, 11.5f); 
                        cb.showText("工艺3  720, 76");  
                        cb.setTextMatrix(780, 9.75f); 
                        cb.showText("工艺4  720, 62");  
                        cb.setTextMatrix(780, 8); 
                        cb.showText("工艺5  720, 48"); 
                        cb.setTextMatrix(780, 6.25f);  
                        cb.showText("标准化   720, 34"); 
                        cb.setTextMatrix(780, 4.5f);  
                        cb.showText("批准   720, 20"); 
                        break;
                    case "A2×3":
                        cb.setFontAndSize(baseFont, 2);  
                        cb.setTextMatrix(728, 35);  //Y轴间隔3.4单位
                        cb.showText("拟制  720, 132");  
                        cb.setTextMatrix(728, 31.8f);  
                        cb.showText("审核  720, 118");  
                        cb.setTextMatrix(728, 28.4f); 
                        cb.showText("工艺1  720, 104");  
                        cb.setTextMatrix(728, 25); 
                        cb.showText("工艺2  720, 90");  
                        cb.setTextMatrix(728, 21.6f); 
                        cb.showText("工艺3  720, 76");  
                        cb.setTextMatrix(728, 18.2f); 
                        cb.showText("工艺4  720, 62");  
                        cb.setTextMatrix(728, 14.8f); 
                        cb.showText("工艺5  720, 48"); 
                        cb.setTextMatrix(728, 11.4f);  
                        cb.showText("标准化   720, 34"); 
                        cb.setTextMatrix(728, 8);  
                        cb.showText("批准   720, 20"); 
                        break;
                    case "A2×5":
                        cb.setFontAndSize(baseFont, 1);  
                        cb.setTextMatrix(772, 21);  //Y轴间隔2单位
                        cb.showText("拟制  772, 21");  
                        cb.setTextMatrix(772, 19);  
                        cb.showText("审核  720, 118");  
                        cb.setTextMatrix(772, 17); 
                        cb.showText("工艺1  720, 104");  
                        cb.setTextMatrix(772, 15); 
                        cb.showText("工艺2  720, 90");  
                        cb.setTextMatrix(772, 13); 
                        cb.showText("工艺3  720, 76");  
                        cb.setTextMatrix(772, 11); 
                        cb.showText("工艺4  720, 62");  
                        cb.setTextMatrix(772, 9); 
                        cb.showText("工艺5  720, 48"); 
                        cb.setTextMatrix(772, 7);  
                        cb.showText("标准化   720, 34"); 
                        cb.setTextMatrix(772, 5);  
                        cb.showText("批准   772, 6"); 
                        break;
                    case "A3×4":
                        cb.setFontAndSize(baseFont, 2);  
                        cb.setTextMatrix(720, 36.5f);  //Y轴间隔3.5单位
                        cb.showText("拟制  720, 21");  
                        cb.setTextMatrix(720, 33);  
                        cb.showText("审核  720, 118");  
                        cb.setTextMatrix(720, 29.5f); 
                        cb.showText("工艺1  720, 104");  
                        cb.setTextMatrix(720, 26); 
                        cb.showText("工艺2  720, 90");  
                        cb.setTextMatrix(720, 22.5f); 
                        cb.showText("工艺3  720, 76");  
                        cb.setTextMatrix(720, 19); 
                        cb.showText("工艺4  720, 62");  
                        cb.setTextMatrix(720, 15.5f); 
                        cb.showText("工艺5  720, 48"); 
                        cb.setTextMatrix(720, 12);  
                        cb.showText("标准化   720, 34"); 
                        cb.setTextMatrix(720, 8.5f);  
                        cb.showText("批准   772, 6"); 
                        break;
                    case "A3×6":
                        cb.setFontAndSize(baseFont, 1.5f);  
                        cb.setTextMatrix(760, 24.3f);  //Y轴间隔2.35单位
                        cb.showText("拟制  760, 21");  
                        cb.setTextMatrix(760, 21.95f);  
                        cb.showText("审核  720, 118");  
                        cb.setTextMatrix(760, 19.6f); 
                        cb.showText("工艺1  720, 104");  
                        cb.setTextMatrix(760, 17.25f); 
                        cb.showText("工艺2  720, 90");  
                        cb.setTextMatrix(760, 14.9f); 
                        cb.showText("工艺3  720, 76");  
                        cb.setTextMatrix(760, 12.55f); 
                        cb.showText("工艺4  720, 62");  
                        cb.setTextMatrix(760, 10.2f); 
                        cb.showText("工艺5  720, 48"); 
                        cb.setTextMatrix(760, 7.85f);  
                        cb.showText("标准化   720, 34"); 
                        cb.setTextMatrix(760, 5.5f);  
                        cb.showText("批准   772, 6"); 
                        break;
                    case "A4×3":
                        cb.setFontAndSize(baseFont, 4);  
                        cb.setTextMatrix(620, 62);  //Y轴间隔6.6单位
                        cb.showText("拟制  520, 132");  
                        cb.setTextMatrix(620, 55.2f);  
                        cb.showText("审核  720, 118");  
                        cb.setTextMatrix(620, 48.6f); 
                        cb.showText("工艺1  720, 104");  
                        cb.setTextMatrix(620, 42); 
                        cb.showText("工艺2  720, 90");  
                        cb.setTextMatrix(620, 35.4f); 
                        cb.showText("工艺3  720, 76");  
                        cb.setTextMatrix(620, 28.8f); 
                        cb.showText("工艺4  720, 62");  
                        cb.setTextMatrix(620, 22.2f); 
                        cb.showText("工艺5  720, 48"); 
                        cb.setTextMatrix(620, 15.6f);  
                        cb.showText("标准化   720, 34"); 
                        cb.setTextMatrix(620, 9);  
                        cb.showText("陶莉 2017/12/28"); 
                        cb.setTextMatrix(765, 34);  
                        cb.showText("版本 F"); 
                        break;
                    case "A4×4":
                        cb.setFontAndSize(baseFont, 3);  
                        cb.setTextMatrix(675, 47);  //Y轴间隔5单位
                        cb.showText("拟制  ");  
                        cb.setTextMatrix(675, 42);  
                        cb.showText("审核  ");  
                        cb.setTextMatrix(675, 37); 
                        cb.showText("工艺1  ");  
                        cb.setTextMatrix(675, 32); 
                        cb.showText("工艺2  ");  
                        cb.setTextMatrix(675, 27); 
                        cb.showText("工艺3  ");  
                        cb.setTextMatrix(675, 22); 
                        cb.showText("工艺4 ");  
                        cb.setTextMatrix(675, 17); 
                        cb.showText("工艺5  "); 
                        cb.setTextMatrix(675, 12);  
                        cb.showText("标准化   "); 
                        cb.setTextMatrix(675, 7);  
                        cb.showText("批准   "); 
                        break;
                    default:
                        break;
                }
                
                 cb.endText(); 
           // }
        }
        catch (Exception e)
        {
            e.printStackTrace();
            throw new Exception(e);
        }finally {
            if (stamper != null) {
                stamper.close();
            }
            if (reader != null) {
                reader.close();
            }
        }
        
    }
    
    
    

    /**  
    * @Title: drawingToPDF 
    * @Description: DWG文件转PDF 
    * @param strFileName //dwg文件名
    * @throws Exception       
    * @return void    返回类型  
    * @throws  
    * @date 2018-11-13
    * @author liuyudong@glaway.com
    */
    public void drawingToPDF(String strFileName ) throws Exception {
        System.out.println(">>>>>>>>>>>>>>drawingToPDF");
            // 读取配置文件
            Properties propertiesUtil = new Properties();
            String ConfigPath = PDFTest.class.getResource("/properties/GW_DrawingToPDFConfig.properties").toString();
            ConfigPath = URLDecoder.decode(ConfigPath, "utf-8");
            propertiesUtil.load(new FileInputStream(ConfigPath.replace("file:/", "")));
            String downloadPath = propertiesUtil.getProperty("downloadPath");//需要被转换的DWG文件目录
            String converterExe = propertiesUtil.getProperty("converterExe");//工具安装目录  如：E:\\Any\\dp.exe 注意目录不能有空格
            String restBat = propertiesUtil.getProperty("restBat");//注册表工具目录 如 E:\\rest_anyDwgtoPDF\\rest_anyDWGtoPDf.bat 不能有空格
            String uploadPath = propertiesUtil.getProperty("uploadPath");//转换为PDF的临时文件夹，转换下载后删除其中文件
            String generatePDFPath = propertiesUtil.getProperty("generatePDFPath");//签名后的pdf保存地址
            // checkout
            File file = new File(downloadPath);
            if(file.exists()){
                String[] strFileNames = file.list();
            
            // converter
             for (String strFileName1:strFileNames) {
               /*  String strFileName = checkoutFileList.getElement(i).getName();
                String converterFileName = strFileName.substring(0, strFileName.indexOf("."))
                        + OtherConstants.SUFFIX_PDF;*/
              //String strFileName ="AL7.825.4047#F_2_01_A3.dwg";
                 if(strFileName1.endsWith(FileConstants.SUFFIX_DWG)){
                 String converterFileName = strFileName1.substring(0,strFileName1.lastIndexOf("."))+FileConstants.SUFFIX_PDF;
              //String converterFileName ="textout.pdf";
                String pdfPath = uploadPath + "\\" + converterFileName;
                String newpdfPath = generatePDFPath + "\\" + converterFileName;
                String command = converterExe + " /InFile " + downloadPath + "\\" + strFileName1 + " /OutFile "
                        + pdfPath;
              
                String format = FileUtil.getInfo(strFileName1);
                if ("A4".equals(format)) {
                    command = command + " /PDFWidth 210 /PDFHeight 297 ";
                } else if ("A3".equals(format)) {
                    command = command + " /PDFWidth 297 /PDFHeight 420 ";
                } else if ("A2".equals(format)) {
                    command = command + " /PDFWidth 420 /PDFHeight 594 ";
                } else if ("A1".equals(format)) {
                    command = command + " /PDFWidth 594 /PDFHeight 841 ";
                } else if ("A0".equals(format)) {
                    command = command + " /PDFWidth 841 /PDFHeight 1189 ";
                } else if ("A1×3".equals(format)) {
                    command = command + " /PDFWidth 841 /PDFHeight 1782 ";
                } else if ("A1×4".equals(format)) {
                    command = command + " /PDFWidth 841 /PDFHeight 2376 ";
                } else if ("A2×3".equals(format)) {
                    command = command + " /PDFWidth 594 /PDFHeight 1260 ";
                } else if ("A2×5".equals(format)) {
                    command = command + " /PDFWidth 594 /PDFHeight 2100 ";
                } else if ("A3×4".equals(format)) {
                    command = command + " /PDFWidth 420 /PDFHeight 1188 ";
                } else if ("A3×6".equals(format)) {
                    command = command + " /PDFWidth 420 /PDFHeight 1782 ";
                } else if ("A4×3".equals(format)) {
                    command = command + " /PDFWidth 297 /PDFHeight 630 ";
                } else if ("A4×4".equals(format)) {
                    command = command + " /PDFWidth 210 /PDFHeight 840 ";
                } 
               
                command = command + "/ConvertType DWG2PDF /IncSubFolder /hide";
                // exec
                CommandUtil.exeCmd(restBat);
                CommandUtil.exeCmd(command);
                //签名
               // generatePDF(pdfPath, newpdfPath);
                //TODO 导出pdf
                
               // FileUtil.deleteDir(uploadPath);//导出后删除pdf
        }
             }
             }
            System.out.println("转换成功");
    }

}


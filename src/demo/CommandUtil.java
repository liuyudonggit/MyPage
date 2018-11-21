package demo;
 
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
 
/**
 * 执行运行时命令
 * @author 李海云
 * @email CloudComputing.cc@gmail.com
 * @date 2012-01-12 15:21
 */
public class CommandUtil {
    /**
     * executeCommand
     * @param command
     * @throws IOException 
     */
    public static void exeCmd(String commandStr){
        BufferedReader br = null;
       try{
           Process p = Runtime.getRuntime().exec(commandStr);
           br = new BufferedReader(new InputStreamReader(p.getInputStream()));
           String line = null;
           StringBuilder sb = new StringBuilder();
           while ((line = br.readLine()) != null)
        {
            sb.append(line + "\\n");
        }
           System.out.println(sb.toString());
       } catch (Exception e) {
           e.printStackTrace();
       }
       finally
       {
           if(br != null)
           {
               try {
                   br.close();
               }catch (Exception e){
                   e.printStackTrace();
               }
           }
       }
    }

    
    public static void main(String[] args)
    {
         String commandStr  =   "E:\\Any\\dp.exe /InFile E:\\DrawingToPDF\\AutoCAD\\DownLoad\\AL7.825.4047#F_2_01_A3.dwg /OutFile E:\\DrawingToPDF\\AutoCAD\\Upload\\textout.pdf /PDFWidth 210 /PDFHeight 297 /ConvertType DWG2PDF /IncSubFolder";
        
         System.out.println( commandStr);
         CommandUtil.exeCmd(commandStr);
    }
}

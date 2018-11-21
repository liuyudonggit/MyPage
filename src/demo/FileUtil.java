package demo;

import java.io.File;

public class FileUtil
{
    
    /**   
    * @ClassName: FileUtil  
    * @Description: TODO  
    * @author liuyudong@glaway.com 
    * @date 2018-11-13 上午11:15:25  
    *  
    *  
    */
    public static void main(String[] args)
    {
        String path = "E:\\DrawingToPDF\\AutoCAD\\Upload";
        deleteDir(path);
    }
    
    //删除文件夹下所有文件
    public static boolean deleteDir(String path)
    {
        File file = new File(path);
        if (!file.exists())
        {// 判断是否待删除目录是否存在
            System.err.println("The dir are not exists!");
            return false;
        }
        
        String[] content = file.list();// 取得当前目录下所有文件和文件夹
        for (String name : content)
        {
            File temp = new File(path, name);
            if (temp.isDirectory())
            {// 判断是否是目录
                deleteDir(temp.getAbsolutePath());// 递归调用，删除目录里的内容
                temp.delete();// 删除空目录
            }
            else
            {
                if (!temp.delete())
                {// 直接删除文件
                    System.err.println("Failed to delete " + name);
                }
            }
        }
        return true;
    }
    
    
    //匹配文档类型
    
    public static String  getInfo(String fileName){
        String format ="";
        fileName = fileName.replace("X", "×");
        if(fileName.contains("×")){
            format = fileName.substring(fileName.lastIndexOf(".")-4,fileName.lastIndexOf("."));
        }else{
            format = fileName.substring(fileName.lastIndexOf(".")-2,fileName.lastIndexOf("."));
        }
        
        return format;
    }
}

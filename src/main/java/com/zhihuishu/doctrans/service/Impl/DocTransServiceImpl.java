package com.zhihuishu.doctrans.service.Impl;

import com.able.base.ftp.oss.OSSPublicUploadInterface;
import com.alibaba.fastjson.JSONObject;
import com.zhihuishu.doctrans.service.DocTransService;
import com.zhihuishu.doctrans.utils.MyFileUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;


import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;


import javax.servlet.http.HttpServletRequest;
import java.io.*;

@Service
public class DocTransServiceImpl implements DocTransService {

    private String datahtml = "/data/html/";

    private String dataimage = "/data/image/";

    private String dataimagesource = "/data/image/word/media/";

    private String datadownload = "/data/download/";


//    @Override
//    public String docx2html(MultipartFile file, String url) {
//        String reshtml = "";
//        try{
//            String str = "";
//            try {
//                File file1 = new File("C:\\Users\\able\\Documents\\Tencent Files\\870070823\\FileRecv\\wordtestpaper2.docx");
//                FileInputStream fis = new FileInputStream(file1);
//                XWPFDocument xdoc = new XWPFDocument(fis);
//                XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
//                String doc1 = extractor.getText();
//                str += doc1;
//                fis.close();
//            } catch (Exception e) {
//                e.printStackTrace();
//            }
//
//        }catch (Exception e){
//            e.printStackTrace();
//        }finally {
//
//        }
//        return reshtml;
//    }

    //String sourceFileName = "C:\\Users\\able\\Documents\\Tencent Files\\870070823\\FileRecv\\wordtestpaper2.docx";
    //http://file.zhihuishu.com/zhs_yufa_150820/ablecommons/demo/201809/3404484743d64189ba968d6328161c9f.docx
    @Override
    public String docx2html(MultipartFile multipartFile, String url, HttpServletRequest request) {
        String str = "";
        String sourceFileName = "";
        File file = null;
        //通过url下载图片
        if(url!=null&&!StringUtils.isEmpty(url)){
            file = MyFileUtil.downloadFile(url,datadownload);
        }
        //文件不存在
        if((file==null||!file.exists())&&multipartFile!=null&&multipartFile.getSize()>0){
            try {
                file = MyFileUtil.inputStreamToFile( multipartFile,request,datadownload);
            }catch (Exception e){
                e.printStackTrace();
                file = null;
            }

        }
        if(file==null){
            return "";
        }
        sourceFileName = datadownload+file.getName();
        String filename = file.getName();
        if(filename!=null&&!StringUtils.isEmpty(filename)){
            filename = filename.split("\\.")[0];
        }
        File filehtmldir = new File(datahtml);
        if (filehtmldir.isDirectory()) {

        }else {
            try {
                filehtmldir.mkdir();
            }catch (Exception e){
                e.printStackTrace();
            }
        }

        String targetFileName = datahtml+filename+".html";
        String imagePathStr = dataimage;
        OutputStreamWriter outputStreamWriter = null;
        try {
            XWPFDocument document = new XWPFDocument(new FileInputStream(sourceFileName));
            XHTMLOptions options = XHTMLOptions.create();
            // 存放图片的文件夹
            options.setExtractor(new FileImageExtractor(new File(imagePathStr)));
            // html中图片的路径
            options.URIResolver(new BasicURIResolver("image"));
            outputStreamWriter = new OutputStreamWriter(new FileOutputStream(targetFileName), "utf-8");
            XHTMLConverter xhtmlConverter = (XHTMLConverter) XHTMLConverter.getInstance();

            xhtmlConverter.convert(document, outputStreamWriter, options);
            //读取html 并返回
            File filehtml = new File(targetFileName);//定义一个file对象，用来初始化FileReader
            FileReader reader = new FileReader(filehtml);//定义一个fileReader对象，用来初始化BufferedReader
            BufferedReader bReader = new BufferedReader(reader);//new一个BufferedReader对象，将文件内容读取到缓存
            StringBuilder sb = new StringBuilder();//定义一个字符串缓存，将字符串存放缓存中
            String s = "";
            while ((s =bReader.readLine()) != null) {//逐行读取文件内容，不读取换行符和末尾的空格
                sb.append(s + "\n");//将读取的字符串添加换行符后累加存放在缓存中
            }
            bReader.close();
            str = sb.toString();
            //遍历文件夹中的图片
            File dataimg = new File(dataimagesource);		//获取其file对象
            File[] fs = dataimg.listFiles();	//遍历path下的文件和目录，放在File数组中
            for(File f:fs){					//遍历File[]数组
                //若非目录(即文件)，则打印
                if(!f.isDirectory()){
                    try{
                        //上传到oss
                        String oss_url = "";
                        oss_url = OSSPublicUploadInterface.ftpAttachment(f,"doctrans","docx2html");
                        if(oss_url!=null&&!StringUtils.isEmpty(oss_url)){
                            JSONObject jsonObject = JSONObject.parseObject(oss_url);
                            JSONObject data = jsonObject.getJSONObject("data");
                            oss_url = data.getString("path");
                            //替换图片链接
                            String img_url = (f+"").replaceAll("\\\\","\\/").replace("/data/","");
                            str = str.replace(img_url,oss_url);
                        }
                        f.delete();
                    }catch (Exception e){
                        e.printStackTrace();
                    }
                }
            }
            //删除本地文件
            if(filehtml!=null){
                filehtml.delete();
            }
            if(file!=null){
                file.delete();
            }
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            if (outputStreamWriter != null) {
                try {
                    outputStreamWriter.close();
                }catch (Exception e){
                    e.printStackTrace();
                }
            }
        }
        return str;
    }




    //            List<XWPFPictureData> listpic = document.getAllPictures();
//            for(XWPFPictureData xwpfPictureData:listpic){
//                System.out.println("imagetype::::::::::::::::"+xwpfPictureData.getPictureType());
//                System.out.println("filename:::::::::::::::::"+xwpfPictureData.getFileName());
//                //矢量图写入
//                if(xwpfPictureData.getPictureType()==3){
//
//                }
//            }


}

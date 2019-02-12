package com.zzzy.controller;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.json.JSONObject;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.client.RestTemplate;

import com.sysware.util.FileUpload;
import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;

@Controller
@RequestMapping("/demo")
public class DemoController {

	
 
	 
    @RequestMapping("/index")
    public String index(){
        return "index";
    }
    @RequestMapping("/stream")
    public String stream(){
    	return "Stream/Stream";
    }
    
    @RequestMapping("/word")
    public String openWord(String fileId,HttpServletRequest request, Map<String,Object> map){
    	String path = request.getServletContext().getRealPath("");
    	
    	PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
		poCtrl.setServerPage(request.getContextPath()+"/poserver.zz");//设置服务页面
		poCtrl.addCustomToolButton("保存","Save()",1);//添加自定义保存按钮
		//poCtrl.addCustomToolButton("盖章","AddSeal",2);//添加自定义盖章按钮
		poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
		//打开word
//		http://test.product.sysware.com.cn/fileserver/download/?ID=201902120953140009538792e8de85a34318a691
		poCtrl.webOpen("wordStream?fileId="+fileId,OpenModeType.docAdmin,"张三");
//		map.put("pageoffice","PageOfficeCtrl1");
		poCtrl.setTagId("PageOfficeCtrl1");
		 
        return "word";
    }
    
    @RequestMapping("/wordStream")
    public void wordStream(String fileId,HttpServletRequest request, 
    		
			HttpServletResponse response) throws IOException {
//		String path ="http://dev.product.sysware.com.cn/fileserver/download/?ID=20190122155122000557d1ae1a09cb72414ca160";
		String path="http://test.product.sysware.com.cn/fileserver/download/?ID="+fileId;
		URL url = new URL(path);
		HttpURLConnection conn = (HttpURLConnection) url.openConnection();
		conn.setConnectTimeout(3 * 1000);
		//防止屏蔽程序抓取而返回403错误
		conn.setRequestProperty("User-Agent", "Mozilla/4.0 (compatible; MSIE 5.0; Windows NT; DigExt)");
	 
		Map<String, List<String>> headerFields = conn.getHeaderFields();
		
		//获取url中的文件名 
		List<String> list = headerFields.get("Content-Disposition");
		String disposition ="";
		String[] split;
		String filename ="default.docx";
		if(list!=null ) {
			 disposition = list.get(0);
		}
		if(disposition.length() >1 && disposition.contains("filename=")) {
			split = disposition.split("filename=");
			String qutoName = split[1];
			if(qutoName.length() >2) {
				 filename = qutoName.substring(1, qutoName.length()-1);
			}
		}
		
		InputStream  inputStream=conn.getInputStream();
	
		 byte[] fileData = readInputStream(inputStream);
		 int fileSize = fileData.length;
		
		response.reset();
		response.setContentType("application/octet-stream"); // application/x-excel, application/ms-powerpoint, application/pdf
		response.setHeader("Content-Disposition",
				"attachment; filename="+filename); //fileN应该是编码后的(utf-8)
		response.setContentLength(fileSize);
		
		OutputStream outputStream = response.getOutputStream();
		
		outputStream.write(fileData);
		
		outputStream.flush();
		outputStream.close();
		outputStream = null;
		
	}
    
    /**
     * 从输入流中获取字节数组
     * @param inputStream
     * @return
     * @throws IOException
     */
    public static  byte[] readInputStream(InputStream inputStream) throws IOException {
        byte[] buffer = new byte[1024];
        int len = 0;
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        while((len = inputStream.read(buffer)) != -1) {
            bos.write(buffer, 0, len);
        }
        bos.close();
        return bos.toByteArray();
    }
    
    
    
    
    @RequestMapping("/word2")
    public String openWord2(HttpServletRequest request, Map<String,Object> map){
    	String path = request.getServletContext().getRealPath("");
    	
    	PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
    	poCtrl.setServerPage(request.getContextPath()+"/poserver.zz");//设置服务页面
    	
    	poCtrl.setCaption("sysware在线编辑");
//    	poCtrl.webOpen("doc/test.doc",OpenModeType.docNormalEdit,"张佚名");
    	poCtrl.webOpen(path+"doc/test.doc",OpenModeType.docAdmin,"张三");
    	poCtrl.setTagId("PageOfficeCtrl1");//此行必需
    	poCtrl.setMenubar(false); 
    	
    	//打开word
    	return "ControlBars/OpenWord";
    }
    
    
    
    @RequestMapping("/testPost")
    public void testPost(HttpServletRequest request1, HttpServletResponse response){
    	  
        String url ="http://127.0.0.1:8080/maven-springmvc/pageoffice/testParam";
        HttpHeaders headers = new HttpHeaders();
        headers.set("phone","123456");
        MultiValueMap<String, Object> params = new LinkedMultiValueMap<>();
        params.add("name","yyc");
        params.add("email","12306");
        params.add("flag","12306");
        RestTemplate restTemplate = new RestTemplate();
        HttpEntity httpEntity = new HttpEntity(params,headers);
        ResponseEntity<String> request = restTemplate.postForEntity(url, httpEntity, String.class);
//        ResponseEntity<String> request = restTemplate.getForEntity(url,String.class,httpEntity,params);
        
    
    }
    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response){
    	String path = request.getServletContext().getRealPath("");
    	String serverUrl ="http://dev.product.sysware.com.cn/fileserver/upload/";
        FileSaver fs = new FileSaver(request, response);
        FileInputStream fileStream = fs.getFileStream();
        String fileName = fs.getFileName();
//        JSONObject upload = FileUpload.upload(serverUrl, fileName, fileStream);
        FileUpload.uploadFileImpl(serverUrl,"F:\\fileTest\\testABC.doc",fileName,null);
        RestTemplate restTemplate = new RestTemplate();
//        restTemplate.postForEntity("http://127.0.0.1:8080/maven-springmvc/pageoffice/testParam", upload, String.class);
        
        String url ="http://127.0.0.1:8080/maven-springmvc/pageoffice/testParam";
        HttpHeaders headers = new HttpHeaders();
//        MediaType type = MediaType.parseMediaType("application/json; charset=UTF-8");
//        headers.setContentType(MediaType.APPLICATION_JSON);
//        headers.add("Accept", MediaType.APPLICATION_JSON.toString());
        MultiValueMap<String, Object> params = new LinkedMultiValueMap<>();
//        Map map =  upload.toMap();
//        params.putAll(map);
//        params.add("result", upload.toString());
        
        HttpEntity httpEntity = new HttpEntity(params, headers);
        restTemplate.postForEntity(url, httpEntity, String.class);
        
//        HttpEntity<Map<String,String>> formEntity = new HttpEntity<Map<String,String>>(map, headers);
//        restTemplate.postForObject(url, formEntity, String.class);
//        restTemplate.getForEntity("http://127.0.0.1:8080/maven-springmvc/pageoffice/testParam", String.class, upload);
        
        fs.saveToFile(path +  fileName);
        fs.close();
    }
    @RequestMapping(value="/testParam",produces = MediaType.APPLICATION_JSON_VALUE)
    public void testParam( String flag,String result,HttpServletRequest request, HttpServletResponse response){
    	System.out.println("A:"+flag);
    	System.out.println("B:"+result);
    }
}
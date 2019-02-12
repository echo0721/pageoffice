package com.sysware.util;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Map;
import java.util.Set;

import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.HttpStatus;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.mime.HttpMultipartMode;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.entity.mime.content.FileBody;
import org.apache.http.entity.mime.content.InputStreamBody;
import org.apache.http.entity.mime.content.StringBody;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.json.JSONObject;

/**
 * Created by Administrator on 2016/1/19.
 */
public class FileUpload {

	 
	public static JSONObject uploadFileImpl(String serverUrl, String localFilePath, String serverFieldName,
			Map<String, String> params) {
		String respStr = null;
		CloseableHttpClient httpclient = HttpClients.createDefault();
		try {
			HttpPost httppost = new HttpPost(serverUrl);
			FileBody binFileBody = new FileBody(new File(localFilePath));
			
			MultipartEntityBuilder multipartEntityBuilder = MultipartEntityBuilder.create().setMode(HttpMultipartMode.RFC6532);
//			multipartEntityBuilder.addPart("file", binFileBody);
			 try {
				multipartEntityBuilder.addBinaryBody("file", binFileBody.getInputStream(), ContentType.DEFAULT_BINARY, serverFieldName);
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
			// 设置上传的其他参数
			setUploadParams(multipartEntityBuilder, params);
			
			HttpEntity reqEntity = multipartEntityBuilder.build();
			httppost.setEntity(reqEntity);
			
			CloseableHttpResponse response=null;
			try {
				response = httpclient.execute(httppost);
				System.out.println(response.getStatusLine());
				HttpEntity resEntity = response.getEntity();
				respStr = getRespString(resEntity);
				EntityUtils.consume(resEntity);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}finally {
				try {
					response.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		} finally {
			
		}
		System.out.println("resp=" + respStr);
		return new JSONObject(respStr);
	}

	private static void setUploadParams(MultipartEntityBuilder multipartEntityBuilder, Map<String, String> params) {
		if (params != null && params.size() > 0) {
			Set<String> keys = params.keySet();
			for (String key : keys) {
				multipartEntityBuilder.addPart(key, new StringBody(params.get(key), ContentType.TEXT_PLAIN));
			}
		}
	}
	 private static String getRespString(HttpEntity entity) throws Exception {
	        if (entity == null) {
	            return null;
	        }
	        InputStream is = entity.getContent();
	        StringBuffer strBuf = new StringBuffer();
	        byte[] buffer = new byte[4096];
	        int r = 0;
	        while ((r = is.read(buffer)) > 0) {
	            strBuf.append(new String(buffer, 0, r, "UTF-8"));
	        }
	        return strBuf.toString();
	    }
	public static JSONObject upload(String url, String fileName, InputStream inputStream) {
		String fdfsPath = "";
		try {
			HttpClient httpclient = new DefaultHttpClient();

			// HttpPost httppost = new
			// HttpPost("http://127.0.0.1:8082/fileUpload.action");//请求格式
			HttpPost httppost = new HttpPost(url);
			// File file = new File(filePath);
			// String name = file.getName();
			// InputStream in = new FileInputStream(file);
			MultipartEntityBuilder reqEntity = MultipartEntityBuilder.create();

			InputStreamBody inputStreamBody = new InputStreamBody(inputStream, fileName);
			StringBody fileNam = new StringBody(fileName);
			// File file = new File("F:\\fileTest\\testABC.doc");
			reqEntity.addBinaryBody(fileName, inputStream);
			reqEntity.addPart("Filename",fileNam);
			// reqEntity.addTextBody("fileName", "testABC.doc");
			HttpEntity build = reqEntity.build();

			httppost.setEntity(build);
			HttpResponse response = httpclient.execute(httppost);
			int statusCode = response.getStatusLine().getStatusCode();

			if (statusCode == HttpStatus.SC_OK) {

				// System.out.println("服务器正常响应.....");
				HttpEntity resEntity = response.getEntity();
				JSONObject json = new JSONObject(EntityUtils.toString(resEntity).toString());

				// System.out.println(EntityUtils.toString(resEntity));//httpclient自带的工具类读取返回数据
				// System.out.println(resEntity.getContent());
				// System.out.println(json);
				EntityUtils.consume(resEntity);
				return json;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}

}
package com.sysware.vo;

public class FileProperty {
	private String username;
	private String fileId;
	private String loginName;
	private String fileName;
	private String openModeType;//只读，编辑，
	private String fileSaveUrl; //文档保存连接，默认文件服务器
	private String paramRecUrl;//文件参数返回方法
	public String getUsername() {
		return username;
	}
	public void setUsername(String username) {
		this.username = username;
	}
	public String getLoginName() {
		return loginName;
	}
	public void setLoginName(String loginName) {
		this.loginName = loginName;
	}
	public String getOpenModeType() {
		return openModeType;
	}
	public void setOpenModeType(String openModeType) {
		this.openModeType = openModeType;
	}
	public String getFileId() {
		return fileId;
	}
	public void setFileId(String fileId) {
		this.fileId = fileId;
	}
	public String getFileSaveUrl() {
		return fileSaveUrl;
	}
	public void setFileSaveUrl(String fileSaveUrl) {
		this.fileSaveUrl = fileSaveUrl;
	}
	public String getParamRecUrl() {
		return paramRecUrl;
	}
	public void setParamRecUrl(String paramRecUrl) {
		this.paramRecUrl = paramRecUrl;
	}
	public String getFileName() {
		return fileName;
	}
	public void setFileName(String fileName) {
		this.fileName = fileName;
	}
	
	
}

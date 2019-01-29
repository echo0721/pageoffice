<%@ page language="java"
	import="java.util.*,com.zhuozhengsoft.pageoffice.*,com.zhuozhengsoft.pageoffice.wordwriter.*"
	pageEncoding="gb2312"%>
<%@ taglib uri="http://java.pageoffice.cn" prefix="po"%>
<%
//******************************卓正PageOffice组件的使用*******************************
	//设置PageOffice服务器组件
// 	PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
// 	poCtrl1.setServerPage(request.getContextPath()+"/poserver.zz"); //此行必须
	
// // 	poCtrl1.setTitlebar(false); //隐藏标题栏
// 	poCtrl1.setMenubar(false); //隐藏菜单栏
// // 	poCtrl1.setOfficeToolbars(false);//隐藏Office工具条
// // 	poCtrl1.setCustomToolbar(false);//隐藏自定义工具栏

// 	//打开文件
// 	poCtrl1.webOpen("doc/template.doc", OpenModeType.docNormalEdit, "张三");

// 	//添加标题
// 	poCtrl1.setCaption("sysware在线编辑");
// 	//添加自定义按钮
// 	poCtrl1.addCustomToolButton("保存","Save()",1);
// 	poCtrl1.setSaveFilePage("SaveFile.jsp");
// // 	poCtrl1.webOpen("doc/template.doc", OpenModeType.docReadOnly, "张三");
// 	poCtrl1.setTagId("PageOfficeCtrl1"); //此行必须	
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
  <head>   
    <title>修改标题栏</title>
    

	
	
  <script type="text/javascript">
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
            //document.getElementById("PageOfficeCtrl1").CustomSaveResult获取的是保存页面的返回值
            document.getElementById("PageOfficeCtrl1").Alert("保存成功，返回值为：" + document.getElementById("PageOfficeCtrl1").CustomSaveResult);
        }
 
    </script>

  </head>
  
  <body>
  	隐藏了标题栏、菜单栏、自定工具栏和Office工具栏的效果，每个栏都是可以单独的控制是否隐藏。
    <div style=" width:auto; height:700px;">
    <po:PageOfficeCtrl id="PageOfficeCtrl1"></po:PageOfficeCtrl>
    </div>
  </body>
</html>

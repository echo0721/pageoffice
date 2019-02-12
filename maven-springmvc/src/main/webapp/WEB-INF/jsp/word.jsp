<%@ page language="java"
	import="java.util.*,com.zhuozhengsoft.pageoffice.*,com.zhuozhengsoft.pageoffice.wordwriter.*"
	pageEncoding="UTF-8"%>
<%@ taglib uri="http://java.pageoffice.cn" prefix="po" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>在线打开文件</title>

  <script type="text/javascript">
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
            //document.getElementById("PageOfficeCtrl1").CustomSaveResult获取的是保存页面的返回值
           // document.getElementById("PageOfficeCtrl1").Alert("保存成功，返回值为：" + document.getElementById("PageOfficeCtrl1").CustomSaveResult);
        }
 
    </script>

</head>
<body>
<div style="width:1000px;height:700px;">
        <po:PageOfficeCtrl id="PageOfficeCtrl1">
        </po:PageOfficeCtrl>
</div>
</body>
</html>
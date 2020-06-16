<%--
  Created by IntelliJ IDEA.
  User: Felix
  Date: 2020/6/16
  Time: 0:12
  To change this template use File | Settings | File Templates.
--%>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
    <title>导出</title>
</head>
<body>
    <button type="button" onclick="exportData()">导出</button>
</body>
</html>
<script type="text/javascript">
    function exportData() {
        window.location = "<%= request.getContextPath() %>/excelExport"
    }
</script>
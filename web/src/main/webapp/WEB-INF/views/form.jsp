<%@ taglib uri="http://java.sun.com/jsp/jstl/core" prefix="c"%>
<%@ taglib uri="http://java.sun.com/jsp/jstl/functions" prefix="fn"%>
<%@ taglib uri="http://www.springframework.org/tags" prefix="s"%>
<%@ taglib uri="http://www.springframework.org/tags/form" prefix="form"%>
<%@ page session="false"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>SIRECA | Crear Usuario</title>
</head>
<body>
	<center>

		<div style="color: teal; font-size: 30px">SIRECA |
			Crear Usuario</div>

		<form:form commandName="userObj" action="${pageContext.request.contextPath}/user/save" method="post">
			<table width="400px" height="150px">
				<tr>
					<td><form:label path="username">Username</form:label></td>
					<td><form:hidden path="id" /><form:input path="username" /></td>
				</tr>
				<tr>
					<td><form:label path="password">Password</form:label></td>
					<td><form:input path="password" /></td>
				</tr>
				<tr>
					<td></td>
					<td><input type="submit" value="Register" />
					</td>
				</tr>
			</table>
		</form:form>


		<a href="<c:url value="/user/list/"/>">Pulsa aquí para ver la lista de usuarios</a>
	</center>
</body>
</html>

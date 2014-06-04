<%@ taglib uri="http://java.sun.com/jsp/jstl/core" prefix="c"%>
<%@ taglib uri="http://java.sun.com/jsp/jstl/functions" prefix="fn"%>
<%@ taglib uri="http://www.springframework.org/tags" prefix="s"%>
<%@ taglib uri="http://www.springframework.org/tags/form" prefix="form"%>
<%@ page session="false"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>SIRECA | Lista de usuarios</title>
</head>
<body>
	<center>

		<div style="color: teal; font-size: 30px">SIRECA | Lista de usuarios</div>

		<c:if test="${!empty userList}">
			<table border="1" bgcolor="black" width="600px">
				<tr
					style="background-color: teal; color: white; text-align: center;"
					height="40px">
					
					<td>Username</td>
					<td>Password</td>
					<td>Edit</td>
					<td>Delete</td>
				</tr>
				<c:forEach items="${userList}" var="user">
					<tr
						style="background-color: white; color: black; text-align: center;"
						height="30px">
						
						<td><c:out value="${user.username}" />
						</td>
						<td><c:out value="${user.password}" />
						</td>
						<td><a href="<c:url value="/user/edit/${user.id}"/>">Edit</a></td>
						<td><a href="<c:url value="/user/delete/${user.id}"/>">Delete</a></td>
					</tr>
				</c:forEach>
			</table>
		</c:if>

		<a href="<c:url value="/user/new"/>">Pulsa aquí para crear un nuevo usuario</a>
	</center>
</body>
</html>

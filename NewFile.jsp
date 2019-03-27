<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"
    import="java.io.*,
    org.apache.poi.poifs.filesystem.POIFSFileSystem,
    org.apache.poi.hssf.record.*,
    org.apache.poi.hssf.model.*,
    org.apache.poi.hssf.usermodel.*,
    org.apache.poi.hssf.util.*" %>
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Insert title here</title>
</head>

<body>
<%
String excelfile="D:\\Excel1.xls";
try{
	POIFSFileSystem fs=new POIFSFileSystem(new FileInputStream(excelfile));
	
	
	//워크북 생성
	HSSFWorkbook workbook=new HSSFWorkbook(fs);
	
	int sheetNum=workbook.getNumberOfSheets();
	
	for(int k=0; k<sheetNum; k++){
		
		//시트 이름과 시트번호를 추출
		%>
		<br>
		Sheet Number <%=k %> <br>
		Sheet Name <%=workbook.getSheetName(k) %> <br>
		<% 
		HSSFSheet sheet=workbook.getSheetAt(k);
		int rows=sheet.getPhysicalNumberOfRows();
		
		for(int r=0; r<rows; r++){
			
			//시트에 대한 행을 하나씩 추출
			HSSFRow row=sheet.getRow(r);
			if(row!=null){
				int cells= row.getPhysicalNumberOfCells();
		%>
				ROW <%=row.getRowNum() %> <%=cells %> <br>
		<% 
			for(short c=0; c<cells; c++){
			
				
				//행에 대한 셀을 하나씩 추출하여 셀 타입에 따라 처리
			HSSFCell cell =row.getCell(c);
			if(cell!=null){
				String value=null;
				
				switch(cell.getCellType()){
				
				case HSSFCell.CELL_TYPE_FORMULA:
					value="FORMULA value=" + cell.getCellFormula();
					break;
				case HSSFCell.CELL_TYPE_NUMERIC:
					value="NUMERIC value=" + cell.getNumericCellValue();
					break;
				case HSSFCell.CELL_TYPE_STRING:
					value="STRING value=" + cell.getStringCellValue();
					break;
				case HSSFCell.CELL_TYPE_BLANK:
					value=null;
					break;
				case HSSFCell.CELL_TYPE_BOOLEAN:
					value="BOOLEAN value=" + cell.getBooleanCellValue();
					break;
				case HSSFCell.CELL_TYPE_ERROR:
					value="ERROR value=" + cell.getErrorCellValue();
					break;
					default:
				}
				%>
				<%="CELL col=" + cell.getCellNum() + "VALUE=" + value %> <br>
				<% 
			}
		}
	}
}
	}
} catch(Exception e){
	%>
	Error occured: <%=e.getMessage() %>
	<% 
	e.printStackTrace();
}

%>
</body>
</html>
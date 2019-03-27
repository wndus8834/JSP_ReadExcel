<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.io.*,
    org.apache.poi.poifs.filesystem.POIFSFileSystem,
    org.apache.poi.hssf.record.*,
    org.apache.poi.hssf.model.*,
    org.apache.poi.hssf.usermodel.*,
    org.apache.poi.hssf.util.*,
    org.apache.poi.poifs.filesystem.*,
    org.apache.poi.ss.usermodel.*,
    org.apache.poi.xssf.usermodel.*" %>
    <%@ page import=" java.sql.Connection,
     java.sql.PreparedStatement,
     java.sql.ResultSet,
     java.util.ArrayList,
     util.DatabaseUtil"
     %>
     <%@ page import="java.util.HashMap, java.util.Set, java.util.Map, java.util.Iterator"%>
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Insert title here</title>
</head>

<body>
<%
String excelfile="D:\\Excel1.xls";
String excelfile2="D:\\excel.xlsx";
HashMap output=new HashMap();
try{
	//POIFSFileSystem fs=new POIFSFileSystem(new FileInputStream(excelfile2));
	
	int rst=-1;
	Connection conn=DatabaseUtil.getConnection();
	PreparedStatement pstmt=null;
	
	
	try{
		//워크북 생성
		XSSFWorkbook workbook=new XSSFWorkbook(new FileInputStream(excelfile2));
		
		int sheetNum=workbook.getNumberOfSheets();
		
		for(int k=0; k<sheetNum; k++){
			
			//시트 이름과 시트번호를 추출
			%>
			<br>
			Sheet Number <%=k %> <br>
			Sheet Name <%=workbook.getSheetName(k) %> <br>
			<% 
			XSSFSheet sheet=workbook.getSheetAt(k);//시트 가져오기
			int rows=sheet.getPhysicalNumberOfRows(); // 행 갯수 가져오기
			String checkid="";
			String checkdate="";
			String checkresult="";
			String description="";
			String equipdivision="";
			String equipid="";
			int checknum=0;
			
			for(int r=0; r<rows; r++){ //row 루프
				
				//시트에 대한 행을 하나씩 추출
				XSSFRow row=sheet.getRow(r);//row 가져오기
				if(row!=null){
					int count=1;
					int cells= row.getPhysicalNumberOfCells();//cell 갯수 가져오기
					StringBuffer sqlBuf=new StringBuffer();
					String sql="INSERT INTO excel VALUES (?,?,?,?,?,?)";
					String a="";
					sqlBuf.append("INSERT INTO excel \n");
					sqlBuf.append("(checkid, checkdate, checkresult, description, equipdivision, equipid) values \n");
					a=sqlBuf.toString();
			%>
					ROW <%=row.getRowNum() %> <%=cells %> <br>
			<% 
					for(short c=0; c<cells; c++){
						//행에 대한 셀을 하나씩 추출하여 셀 타입에 따라 처리
						XSSFCell cell =row.getCell(c);//cell 가져오기

						String value=null;
						if(cell!=null){
							
							switch(cell.getCellType()){//cell 타입에 따른 데이타 저장
							
							case XSSFCell.CELL_TYPE_FORMULA:
								value="FORMULA value=" + cell.getCellFormula();
								break;
							case XSSFCell.CELL_TYPE_NUMERIC:
								value="NUMERIC value=" + cell.getNumericCellValue();
								break;
							case XSSFCell.CELL_TYPE_STRING:
								value="STRING value=" + cell.getStringCellValue();
								break;
							case XSSFCell.CELL_TYPE_BLANK:
								value=null;
								break;
							case XSSFCell.CELL_TYPE_BOOLEAN:
								value="BOOLEAN value=" + cell.getBooleanCellValue();
								break;
							case XSSFCell.CELL_TYPE_ERROR:
								value="ERROR value=" + cell.getErrorCellValue();
								break;
								default:
							}
							
							if(row.getRowNum()==2 && cell.getColumnIndex()==8 ){
								equipid=value;
							}
							%>
							<%="CELL col="+ cell.getColumnIndex() + "VALUE=" + value %> <br>
							<% 
							pstmt = conn.prepareStatement(a);
							pstmt.setString(count, value);
							count++;
						}

		
					}
				}
			}
		}
		output.put("msgrst", "추가되었습니다.");
		output.put("result", "1");
		try {if(conn!=null) conn.close();} catch(Exception e) {e.printStackTrace();}
	}catch(Exception e){
		
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
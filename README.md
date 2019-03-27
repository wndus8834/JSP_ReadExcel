## JSP_ReadExcel

poi를 이용하여 엑셀 파일 읽기

![image](https://user-images.githubusercontent.com/38427658/55107039-b0908300-5113-11e9-83a3-364cdb37c349.png)

![image](https://user-images.githubusercontent.com/38427658/55106971-8048e480-5113-11e9-9796-f78e005164a9.png)

1. 엑셀파일 경로를 불러옵니다.
```java
String excelfile2="D:\\Excel1.xlsx";
```

2. 엑셀파일을 로드합니다. (워크북 생성)
```java
XSSFWorkbook workbook=new XSSFWorkbook(new FileInputStream(excelfile2));
```

3. 시트갯수를 가져옵니다.
```java
int sheetNum=workbook.getNumberOfSheets();
```

4. for문을 이용하여 시트 이름과 시트번호를 추출합니다.
```java
for(int k=0; k<sheetNum; k++){
		
		//시트 이름과 시트번호를 추출
		%>
		<br>
		Sheet Number <%=k %> <br>
		Sheet Name <%=workbook.getSheetName(k) %> <br>
}
```

5. 시트에 대한 행을 하나씩 추출합니다.
```java
XSSFCell cell =row.getCell(c);
```

6. 행에 대한 셀을 하나씩 추출하여 셀 타입에 따라 처리합니다.
```java
switch(cell.getCellType()){
				
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
```
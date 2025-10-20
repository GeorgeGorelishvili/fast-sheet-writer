# ექსელის ექსპორტის მოდული #

### **ექსელის რეპორტის დასაგენერირებლად გამარტივებული API.** ###

## გამოყენების ინსტრუქცია ##

 ## ინსტანსის შექმნა FastSheetWriterImpl ##

```
#!java

		FastSheetWriter fsw = FastSheetWriterImpl.getInstance();
```

ავტომატურად დამატებულია ექსელის უჯრების ფორმატები: DATE, DATETIME, AMOUNT და ALLOW_WRAP; შემდეგი ფორმატებით და სიტყვა-გასაღებებით:

```
#!java
		public class Format {
			public static final String AMOUNT 	= "0.00";
			public static final String DATE 	= "dd/MM/yyyy";
			public static final String DATETIME = "dd/MM/yyyy hh:mm";
		}
```


```
#!java
		public class Keys {
			public static final String AMOUNT = "amount";
			public static final String DATE = "date";
			public static final String DATETIME = "datetime";
			public static final String ALLOW_WRAP = "allow_wrap";
		}
```


## ახალი სტილის დამატება ##
```
#!java
		XSSFCellStyle style = fsw.getWorkbook().createCellStyle();
		style.setAlignment(HorizontalAlignment.RIGHT);
		XSSFDataFormat fmt = fsw.getWorkbook().createDataFormat();
		...
		fsw.addStyle("KEY", style);
```
## ან არსებული სტილის გადაფარვა ##
```
#!java
		XSSFCellStyle amountStyle = fsw.getWorkbook().createCellStyle();
		amountStyle.setAlignment(HorizontalAlignment.RIGHT);
		XSSFDataFormat fmt = fsw.getWorkbook().createDataFormat();
		...
		fsw.addStyle(Keys.AMOUNT, amountStyle);
```

ექსელის ჩანართის (Sheet) შექმნა:  `fsw.createSheet();` ან `fsw.createSheet("sheet_name");`
სადაც "sheet_name" არის ჩანართის სახელი. უპარამეტრო მეთოდის შემთხვევაში არის "sheet".

## ექსელის ცხრილის უჯრის და ხაზის დამატების API  ## 

```
#!java
		void addRow() throws IOException;

		void endRow() throws IOException;

		void addEmptyRow() throws IOException;

		void addCell(Object value) throws IOException;

		void addCell(Object value, String styleKey) throws IOException;

		void addFormulaCell(String formula) throws IOException;

		void addFormulaCell(String formula, String styleKey) throws IOException;

		void addFormulaCell(int columnIndex, String formula, String styleKey) throws IOException;

		void addFirstCell(Object value) throws IOException;

		void addFirstCell(Object value, String styleKey) throws IOException;

		void addNewRowCell(Object value) throws IOException;

		void addNewRowCell(Object value, String styleKey) throws IOException;

		void addLastCell(Object value) throws IOException;

		void addLastCell(Object value, String styleKey) throws IOException;
```

## ნებისმიერი `Cell`-ის დამატებისას API თვითონ ზრუნავს და ასწორებს ლოგიკას რომ ვალიდური ექსელის ცხრილი დაგენერირდეს. ##
 მაგ: 

```
#!java
		FastSheetWriter fsw = FastSheetWriterImpl.getInstance(); // line 1
		fsw.addCell("text"); // line 2
		fsw.addCell(amount_value, Keys.AMOUNT); // line 3
		fsw.saveReport("report_path", "file_name"); // line 4
```
 აკეთებს შემდეგს: 

1.  ქმნის ექსელის ფაილს (line 1)
2. აინიციალიზებს სტილებს თარიღებისთვის, თანხისთვის და ახალი ხაზის მხარდაჭერისთვის. (line 1)
3.  ქმნის გაჩუმებითი სახელით ჩანართს (line 2) 
4.  ქმნის ახალ ხაზს (line 2) 
5.  ამატებს ექსელის უჯრას, სადაც წერს amount_value ცვლადის მნიშვნელობას. (line 3) 
6.  აფორმატირებს როგორც თანხის ველი (line 3) 
7.  ამთავრებს ექსელის ხაზის ჩაწერას (line 4) 
8.  ამთავრებს ექსელის ჩანართის ჩაწერას (line 4) 
9.  ქმნის "report_path" დირექტორიას არ არსებობის შემთხვევაში (line 4) 
10.  ქმნის ექსელის ფაილს "report_path" დირექტორიაში "file_name" სახელით (line 4). 

გაქვს ასევე API ფუნქციები გავიგოთ თუ სად ჩაწერს `addCell(value)` მეთოდის გამოძახებისას "value" მნიშვნელობას

```
#!java
		int getCurrentRowIndex();
		int getCurrentColumnIndex();
```

## ხაზის ან სვეტის გამოტოვება ##

```
#!java
		void setStartRowIndex(int startRowIndex);
		void setStartColumnIndex(int startColumnIndex);
```

## წინასწარ მომზადებული მონაცემების ერთიანად ჩაწერა ##
```
#!java
		addData(List<ColumnHeader> headers, List<Object[]> data)
```
სადაც 
```
#!java
		public class ColumnHeader {
			private String name; // ველის სათაური
			private String styleKey; // წინასწარ დამატებული სტილის სიტყვა-გასაღები
			...
```
თუმცა Date, Float, Double ავტომატურად ფორმატირდება შესაბამისად "DATE" და "AMOUNT" ფორმატებით როდესაც styleKey ცარიელია ან შესაბამისი ფორმატის სტილები გადაფარულია.
/*
package com.easy.convert.util;

import com.aspose.words.Font;
import com.aspose.words.*;
import org.apache.commons.lang.StringUtils;

import javax.imageio.ImageIO;
import java.awt.*;
import java.io.File;
import java.util.List;
import java.util.*;
import java.util.regex.Pattern;

public class AsposeWordUtil {

	public static String	GRAMMAR_PREFIX	= "${";									// 语法前缀
	public static String	GRAMMAR_SUFFIX	= "}";									// 后缀

	public static String	textSymbol		= "";
	public static String	imageSymbol		= "@";
	public static String	tableSymbol		= "#";

	public static String    underLineSymbol = "-";

	public static String	prefixRegex		= escapeExprSpecialWord(GRAMMAR_PREFIX);
	public static String	suffixRegex		= escapeExprSpecialWord(GRAMMAR_SUFFIX);

	*/
/**
	 * 模板渲染(包括文字，图片，表格)
	 * 
	 * @param templatePath
	 * @param data2
	 * @param outputPath
	 * @throws Exception
	 *//*

	public static void render(String templatePath, Map<String, Object> data, String outputPath,Boolean flag) throws Exception {

		Document doc = new Document(templatePath);

		String[] runs = doc.getRange().getText().split(escapeExprSpecialWord(GRAMMAR_PREFIX));
		for (String run : runs) {
			if (run.indexOf(GRAMMAR_SUFFIX) != -1) {
				String name = run.substring(0, run.indexOf(GRAMMAR_SUFFIX));
				if (name.startsWith(imageSymbol)) {
					renderImage(doc, data, name.substring(imageSymbol.length()));
				} else if (name.startsWith(tableSymbol)) {
					renderTable(doc, data, name.substring(tableSymbol.length()));
				} else if (name.startsWith(underLineSymbol)){
					renderText(doc, data, name.substring(underLineSymbol.length()),flag,true);
				} else {
					renderText(doc, data, name.substring(textSymbol.length()),flag,false);
				}
			} else {
				continue;
			}
		}
		doc.save(outputPath);
	}

	*/
/**
	 * 只渲染图片
	 * 
	 * @param doc
	 * @param data
	 *//*

	public static void renderImage(Document doc, Map<String, Object> data, String name) throws Exception {

		Object value = data.get(name);
		if (value == null || !(value instanceof ImageData)) {
			doc.getRange().replace(Pattern.compile(prefixRegex + "\\" + imageSymbol + escapeExprSpecialWord(name) + suffixRegex), "");
			return;
		}

		ImageData imageData = (ImageData) value;
		doc.getRange().replace(Pattern.compile(prefixRegex + "\\" + imageSymbol + escapeExprSpecialWord(name) + suffixRegex), new IReplacingCallback() {
			@Override
			public int replacing(ReplacingArgs e) throws Exception {
				DocumentBuilder builder = new DocumentBuilder((Document) e.getMatchNode().getDocument());
				builder.moveTo(e.getMatchNode());

				builder.insertImage(ImageIO.read(new File(imageData.getPath())), imageData.getWidth(), imageData.getHeight());

				e.setReplacement("");
				return ReplaceAction.REPLACE;
			}
		}, false);

	}

	*/
/**
	 * 只渲染文字
	 *
	 * @param doc
	 * @param data
	 * @throws Exception
	 *//*

	public static void renderText(Document doc, Map<String, Object> data, String name,Boolean flag, Boolean isUnder) throws Exception {

		Object value = data.get(name);
		if (value == null) {
			doc.getRange().replace(Pattern.compile(prefixRegex + escapeExprSpecialWord(name) + suffixRegex), "");
			return;
		}
		if (isUnder){
			name = underLineSymbol + name;
		}
		doc.getRange().replace(Pattern.compile(prefixRegex + escapeExprSpecialWord(name) + suffixRegex), new IReplacingCallback() {
			@Override
			public int replacing(ReplacingArgs e) throws Exception {
				DocumentBuilder builder = new DocumentBuilder((Document) e.getMatchNode().getDocument());
				builder.moveTo(e.getMatchNode());
				Font font = builder.getFont();
				if (flag){
					//设置文字颜色
//					font.setColor(Color.RED);
					//设置文字背景色
					font.setHighlightColor(Color.yellow);

				}
				//文字下划线
				if (isUnder){
					font.setUnderline(1);
				}


				builder.write(value.toString());
				e.setReplacement("");
				return ReplaceAction.REPLACE;
			}
		}, false);
	}


	*/
/**
	 * 只渲染表格
	 * 
	 * @param doc
	 * @param data
	 * @throws Exception
	 *//*

	public static void renderTable(Document doc, Map<String, Object> data, String name) throws Exception {

		Object value = data.get(name);
		if (value == null || !(value instanceof TableData)) {
			doc.getRange().replace(Pattern.compile(prefixRegex + "\\" + tableSymbol + escapeExprSpecialWord(name) + suffixRegex), "");
			return;
		}

		TableData tableData = (TableData) value;
		doc.getRange().replace(Pattern.compile(prefixRegex + "\\" + tableSymbol + escapeExprSpecialWord(name) + suffixRegex), new IReplacingCallback() {
			@Override
			public int replacing(ReplacingArgs e) throws Exception {
				DocumentBuilder builder = new DocumentBuilder((Document) e.getMatchNode().getDocument());
				builder.moveTo(e.getMatchNode());

				Table table = builder.startTable();
				builder.setBold(true);
				table.setBorders(LineStyle.SINGLE, tableData.getBorder(), Color.BLACK);

				List<String> headers = tableData.getHeaders();
				for (String header : headers) {
					builder.insertCell();
					builder.write(header);
					builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
					builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
				}
				builder.endRow();

				builder.setBold(false);

				List<Row> rows = tableData.getRows();
				for (Row row : rows) {
					for (String td : row.getRowData()) {
						builder.insertCell();
						builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
						builder.getParagraphFormat().setAlignment(ParagraphAlignment.LEFT);
						builder.write(td);
					}
					builder.endRow();
				}
				builder.endTable();

				e.setReplacement("");
				return ReplaceAction.REPLACE;
			}
		}, false);
	}

	*/
/**
	 * 转换格式
	 * 
	 * @param inputFilePath
	 * @param outputFilePath
	 * @param saveFormat
	 * @throws Exception
	 *//*


	public static List<String> convertFormat(String filePath, String saveFormat, String outputFilePath) throws Exception {
		if (saveFormat == null || saveFormat.isEmpty()) {
			throw new IllegalArgumentException("save format can not be empty..");
		}

		int format;

		if (saveFormat.equalsIgnoreCase("doc")) {
			format = SaveFormat.DOC;
		} else if (saveFormat.equalsIgnoreCase("docx")) {
			format = SaveFormat.DOCX;
		} else if (saveFormat.equalsIgnoreCase("pdf")) {
			format = SaveFormat.PDF;
		} else if (saveFormat.equalsIgnoreCase("html")) {
			format = SaveFormat.HTML;
		} else if (saveFormat.equalsIgnoreCase("text")) {
			format = SaveFormat.TEXT;
		} else if (saveFormat.equalsIgnoreCase("jpeg")) {
			Document doc = new Document(filePath);
			ImageSaveOptions iso = new ImageSaveOptions(SaveFormat.JPEG);
			iso.setResolution(100);
			iso.setPrettyFormat(true);
			iso.setUseAntiAliasing(true);

			List<String> imagesPath = new ArrayList<>();
			String fileName = "image_" + UUID.randomUUID();
			for (int i = 0; i < doc.getPageCount(); i++) {
				iso.setPageIndex(i);
				doc.save(outputFilePath + fileName + i+".jpeg", iso);
				imagesPath.add(fileName + i+".jpeg");
			}


			return imagesPath;
		} else {
			throw new IllegalArgumentException("unknown save format:" + saveFormat);
		}
		Document doc = new Document(filePath);
		doc.save(outputFilePath, format);
		return Arrays.asList(outputFilePath);
	}

	*/
/**
	 * 替换特殊字符
	 * 
	 * @param keyword
	 * @return
	 *//*

	public static String escapeExprSpecialWord(String keyword) {
		if (!(keyword == null || keyword.equals(""))) {
			String[] fbsArr = { "\\", "$", "(", ")", "*", "+", ".", "[", "]", "?", "^", "{", "}", "|" };
			for (String key : fbsArr) {
				if (keyword.contains(key)) {
					keyword = keyword.replace(key, "\\" + key);
				}
			}
		}
		return keyword;
	}

	*/
/**
	 * 获取所有的关键字
	 * @param filePath
	 * @return
	 * @throws Exception
	 *//*

	public static String findAllExprSpecialWord(String filePath) throws Exception{
		Document doc = new Document(filePath);
		StringBuffer sb = new StringBuffer();
		String[] runs = doc.getRange().getText().split(escapeExprSpecialWord(GRAMMAR_PREFIX));
		for (String run : runs) {
			if (run.indexOf(GRAMMAR_SUFFIX) != -1) {
				String name = run.substring(0, run.indexOf(GRAMMAR_SUFFIX));
				sb.append("${" + name + "},");
			} else {
				continue;
			}
		}
		if (sb.length() > 0) {
			String exprSpecialWord = sb.toString().substring(0,sb.toString().length() -1);
			return exprSpecialWord;
		}
		return sb.toString();
	}

	*/
/***
	 * 把doc中内容追加到第一个doc中
	 * @param wordPath 文件路径的集合
	 * @param outPath 输出路径
	 * @return
	 *//*

	public static void appendDoc(List<String> wordPath, String outPath) throws Exception {
		// 判断是否大于一个word
		if (wordPath == null || wordPath.isEmpty() || wordPath.size()<=1){
			return ;
		}


		// 把所有的doc找出来放到一个集合中来
		List<Document> docs = new ArrayList<>();
		for (String s : wordPath) {
			Document document = new Document(s);
			docs.add(document);
		}

		// 循环都追加到第一个doc中去
		Document firstDoc = docs.get(0);

		for (int i = 1; i < docs.size(); i++) {
			firstDoc.appendDocument(docs.get(i), ImportFormatMode.USE_DESTINATION_STYLES);
		}

		firstDoc.save(outPath, SaveFormat.DOCX);
	}

	@SuppressWarnings("serial")
	public static void main(String[] args) throws Exception {

		String path = "C:\\Users\\Administrator\\Desktop\\test.docx";
		String outPath = "C:\\Users\\Administrator\\Desktop\\test-out.docx";
		String imageOutPath = "C:\\Users\\Administrator\\Desktop\\";

		//关键字的map
		List<ExprSpecialWords> exprSpecialWordsList = new ArrayList<>();
		List<String> keyWordsList = new ArrayList<>();
		//word的图片路径
		List<String> wordImageUrls = new ArrayList<String>();

		//提取关键字
		String keyWords = AsposeWordUtil.findAllExprSpecialWord(path);
		if (org.apache.commons.lang3.StringUtils.isNotEmpty(keyWords)){
			String[] exprSpecialWords = keyWords.split(",");
			for (int i = 0; i < exprSpecialWords.length; i++) {
				ExprSpecialWords words = new ExprSpecialWords();
				words.setId(i);
				words.setName(exprSpecialWords[i]);
				exprSpecialWordsList.add(words);
				keyWordsList.add(exprSpecialWords[i]);
			}
		}

		// 把关键字都替换成空格占位
		Map<String,Object> dataMap = new HashMap<>();
		for (String s : keyWordsList) {
			s = s.replaceAll("[-]","").replace("${","").replace("}","");
			String dataStr = "               ";
			dataMap.put(s,dataStr);
		}

		AsposeWordUtil.render(path,dataMap,outPath,false);

		// 图片导出路径 url为空白关键字的word
//		List<String> imageUrls = AsposeWordUtil.convertFormat(outPath,"jpeg",imageOutPath);
	}

	public static class TableData {
		private List<String>	headers;
		private List<Row>		rows;
		private double			width;
		private double			border;

		public List<String> getHeaders() {
			return headers;
		}

		public void setHeaders(List<String> headers) {
			this.headers = headers;
		}

		public List<Row> getRows() {
			return rows;
		}

		public void setRows(List<Row> rows) {
			this.rows = rows;
		}

		public double getWidth() {
			return width;
		}

		public void setWidth(double width) {
			this.width = width;
		}

		public double getBorder() {
			return border;
		}

		public void setBorder(double border) {
			this.border = border;
		}

		public TableData(List<String> headers, List<Row> rows, double width, double border) {
			super();
			this.headers = headers;
			this.rows = rows;
			this.width = width;
			this.border = border;
		}

		public TableData(List<String> headers, List<Row> rows) {
			super();
			this.headers = headers;
			this.rows = rows;
		}

	}

	public static class Row {
		private List<String> rowData;

		public List<String> getRowData() {
			return rowData;
		}

		public void setRowData(List<String> rowData) {
			this.rowData = rowData;
		}

		public Row(List<String> rowData) {
			super();
			this.rowData = rowData;
		}

	}

	public static class ImageData {
		private String	path;
		private double	height;
		private double	width;

		public String getPath() {
			return path;
		}

		public void setPath(String path) {
			this.path = path;
		}

		public double getHeight() {
			return height;
		}

		public void setHeight(double height) {
			this.height = height;
		}

		public double getWidth() {
			return width;
		}

		public void setWidth(double width) {
			this.width = width;
		}

		public ImageData(String path, double height, double width) {
			super();
			this.path = path;
			this.height = height;
			this.width = width;
		}

	}
}
*/

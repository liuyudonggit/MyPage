package ext.print;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Pattern;

import matrix.db.Context;
import matrix.db.FileList;
import matrix.util.StringList;

import org.apache.commons.lang3.StringUtils;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Image;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfGState;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.matrixone.apps.domain.DomainConstants;
import com.matrixone.apps.domain.DomainObject;
import com.matrixone.apps.domain.util.ContextUtil;
import com.matrixone.apps.domain.util.EnoviaResourceBundle;
import com.matrixone.apps.domain.util.FrameworkException;
import com.matrixone.apps.domain.util.MapList;
import com.matrixone.apps.framework.ui.UIUtil;

import ext.constants.AttributeConstants;
import ext.constants.OtherConstants;
import ext.constants.PolicyConstants;
import ext.constants.RelationShipConstants;
import ext.constants.TypeConstants;
import ext.util.FileUtil;
import ext.util.PersonUtil;
import ext.util.Util;

@SuppressWarnings({ "rawtypes", "deprecation", "unchecked" })
public class PrintDocuments {
	private boolean isUsing = false;
	private static PrintDocuments print;
	private SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
	private List<DomainObject> allDocumentList;
	private List<Map> allDocumentInfoList;
	private List<DomainObject> routeList;
	private List<Map> allPrintSettingList;
	private List<MapList> allPrintSetParamsList;
	private String app_path;
	private String temp_path;
	private Context context;
	private static final int wdFormatPDF = 17;
	private Map<String, String> stamperMap;
	private int attachmentCount;
	private boolean isConvertUsing = false;

	private Map<String, String> returnMap = new HashMap<String, String>();

	private PrintDocuments() {
	}

	public static PrintDocuments getInstance() {
		return (print == null) ? new PrintDocuments() : print;
	}

	/**
	 * insert electronic signature and approval info to documents
	 * 
	 * @param context
	 *            The Matrix Context.
	 * @param args
	 *            holds input arguments.
	 * @return Map
	 * @throws Exception
	 *             If the operation fails.
	 */

	public Map<String, String> print(Context context, List<String> docIdList, String app_path, String temp_path,
			Map<String, String> map) {
		try {
			System.out.println("start print................");
			ContextUtil.pushContext(context, "Test Everything", "", "");
			// 校验当前打印是否在使用中
			if (isUsing) {
				Map<String, String> returnMap = new HashMap<String, String>();
				getErrorInfo(context, returnMap, "", "emxEngineeringCentral.Print.Error");
				return returnMap;
			} else {
				isUsing = true;
			}
			System.out.println("init data.............");
			// 始化数据
			init(context, app_path, temp_path, map);
			System.out.println("check data.............");
			// 打印校验
			check(docIdList);

			System.out.println("print Document............." + returnMap);
			// 打印文档
			if (OtherConstants.FAIL.equals(returnMap.get(OtherConstants.ReturnResult))) {
				return returnMap;
			}
			printDocument();// error

			ContextUtil.popContext(context);
			returnMap.put(OtherConstants.ReturnResult, OtherConstants.SUCCESS);
			System.out.println("end print................");
		} catch (Exception e) {
			e.printStackTrace();
			getErrorInfo(context, returnMap, "", "emxEngineeringCentral.Print.Error1");
		} finally {
			if (print != null) {
				print = null;
			}
		}
		return returnMap;
	}

	private void init(Context context, String app_path, String temp_path, Map<String, String> map) {
		this.context = context;
		this.app_path = app_path;
		this.temp_path = temp_path;
		this.stamperMap = map;
		allDocumentList = new ArrayList<DomainObject>();
		allDocumentInfoList = new ArrayList<Map>();
		routeList = new ArrayList<DomainObject>();
		allPrintSettingList = new ArrayList<Map>();
		allPrintSetParamsList = new ArrayList<MapList>();
	}

	private void check(List<String> docIdList) throws Exception {
		for (String documentId : docIdList) {
			// 校验流程是否存在
			DomainObject document = DomainObject.newInstance(context, documentId);
			StringList stringList = new StringList(DomainConstants.SELECT_TYPE);
			stringList.add(DomainConstants.SELECT_NAME);
			stringList.add(AttributeConstants.Select_GW_Format);
			stringList.add(DomainConstants.SELECT_REVISION);
			Map documentInfoMap = document.getInfo(context, stringList);
			String name = (String) documentInfoMap.get(DomainConstants.SELECT_NAME);
			String version = (String) documentInfoMap.get(DomainConstants.SELECT_REVISION);

			String strRouteId = getRelatedRouteId(document);
			DomainObject route = null;
			if (UIUtil.isNotNullAndNotEmpty(strRouteId)) {
				route = DomainObject.newInstance(context, strRouteId);
			}
			System.out.println("check data : strRouteId : " + strRouteId);
			System.out.println("check data : documentInfoMap : " + documentInfoMap);
			if ("Archives".equals(stamperMap.get("printAction")) && route == null) {
				getErrorInfo(context, returnMap, name + " " + version, "emxEngineeringCentral.Print.Error2");
				return;
			} else {
				routeList.add(route);
				allDocumentInfoList.add(documentInfoMap);
				allDocumentList.add(document);
			}
			// 校验打印配置是否存在
			String strDocType = (String) documentInfoMap.get(DomainConstants.SELECT_TYPE);
			String strSizeFormat = (String) documentInfoMap.get(AttributeConstants.Select_GW_Format);
			System.out.println("check data  : strDocType : " + strDocType);
			System.out.println("check data  : strSizeFormat : " + strSizeFormat);
			MapList printSettingMapList = getPrintSetting(strDocType, strSizeFormat);
			System.out.println("check data  : printSettingMapList : " + printSettingMapList);
			if (Util.isEmptyList(printSettingMapList)) {
				getErrorInfo(context, returnMap, name + " " + version, "emxEngineeringCentral.Print.Error3");
				return;
			} else {
				Map printSettingMap = (Map) printSettingMapList.get(0);
				String printSettingId = (String) printSettingMap.get(DomainConstants.SELECT_ID);
				MapList printSetParamsMapList = getPrintSettingParams(printSettingId);
				System.out.println("check data : printSetParamsMapList : " + printSetParamsMapList);
				if (Util.isEmptyList(printSetParamsMapList)) {
					getErrorInfo(context, returnMap, name + " " + version, "emxEngineeringCentral.Print.Error3");
					return;
				} else {
					allPrintSettingList.add(printSettingMap);
					allPrintSetParamsList.add(printSetParamsMapList);
				}
			}
		}
	}

	/**
	 * get related route id
	 * 
	 * @param context
	 *            The Matrix Context.
	 * @param DomainObject
	 *            context object
	 * @return String
	 * @throws Exception
	 *             If the operation fails.
	 */
	private String getRelatedRouteId(DomainObject document) throws Exception {
		String strRouteId = "";
		if (document.isKindOf(context, TypeConstants.GW_DesignDoc)) {
			strRouteId = (String) document.getInfo(context, "from[" + DomainConstants.RELATIONSHIP_OBJECT_ROUTE
					+ "].to.id");
			if (UIUtil.isNullOrEmpty(strRouteId)) {
				String strDDCOId = (String) document.getInfo(context, "to[" + RelationShipConstants.GW_AffectedItems
						+ "].from.id");
				if (UIUtil.isNotNullAndNotEmpty(strDDCOId)) {
					DomainObject domDDCO = DomainObject.newInstance(context, strDDCOId);
					strRouteId = (String) domDDCO.getInfo(context, "from[" + DomainConstants.RELATIONSHIP_OBJECT_ROUTE
							+ "].to.id");
				} else {
					String partApprovalRequestId = (String) document.getInfo(context, "to["
							+ RelationShipConstants.GW_PartApprovalItems + "].from.id");
					if (UIUtil.isNotNullAndNotEmpty(partApprovalRequestId)) {
						DomainObject partApprovalRequest = DomainObject.newInstance(context, partApprovalRequestId);
						strRouteId = (String) partApprovalRequest.getInfo(context, "from["
								+ DomainConstants.RELATIONSHIP_OBJECT_ROUTE + "].to.id");
					}
				}
			}
		} else {
			strRouteId = (String) document.getInfo(context, "from[" + DomainConstants.RELATIONSHIP_OBJECT_ROUTE
					+ "].to.id");
		}
		return strRouteId;
	}

	/**
	 * get print setting id
	 * 
	 * @param context
	 *            The Matrix Context.
	 * @param string
	 *            setting name
	 * @param String
	 *            setting revision
	 * @param String
	 *            setting parameters
	 * @return String
	 * @throws Exception
	 *             If the operation fails.
	 */
	private MapList getPrintSetting(String strDocType, String strDocFormat) throws Exception {

		StringBuffer where = new StringBuffer();

		where.append(AttributeConstants.Select_GW_PrintDocType + " == '" + strDocType + "'");
		if (UIUtil.isNotNullAndNotEmpty(strDocFormat))
			where.append(" && " + AttributeConstants.Select_GW_Format + " == '" + strDocFormat + "'");

		StringList busSelects = new StringList(DomainConstants.SELECT_ID);
		busSelects.add(AttributeConstants.Select_GW_PrintPageNumber);

		return DomainObject.findObjects(context, TypeConstants.GW_PrintSetting, "*", "*", "*", "*", where.toString(),
				false, busSelects);
	}

	private MapList getPrintSettingParams(String printSettingId) throws Exception {
		DomainObject printSetting = DomainObject.newInstance(context, printSettingId);

		StringList busSelects = new StringList(DomainConstants.SELECT_ID);
		busSelects.addElement(AttributeConstants.Select_GW_PrintImgHeight);
		busSelects.addElement(AttributeConstants.Select_GW_PrintImgWidth);
		busSelects.addElement(AttributeConstants.Select_GW_PrintPosition_x);
		busSelects.addElement(AttributeConstants.Select_GW_PrintPosition_y);
		busSelects.addElement(AttributeConstants.Select_GW_PrintIsImg);
		busSelects.addElement(AttributeConstants.Select_GW_PrintSequence);
		busSelects.addElement(AttributeConstants.Select_GW_PrintTag);
		busSelects.addElement(AttributeConstants.Select_GW_PrintFontSize);
		busSelects.addElement(AttributeConstants.Select_GW_PrintPageNumber);
		busSelects.addElement(AttributeConstants.Select_GW_PrintRotation);
		busSelects.addElement(AttributeConstants.Select_GW_PrintOpacity);
		busSelects.addElement(AttributeConstants.Select_GW_PrintStamperWidth);

		return printSetting.getRelatedObjects(context, RelationShipConstants.GW_PrintSettingParams,
				TypeConstants.GW_PrintSettingParam, busSelects, null, false, true, (short) 1, "", null, (short) 0);
	}

	private void printDocument() throws Exception {
		for (int i = 0; i < allDocumentList.size(); i++) {
			DomainObject document = allDocumentList.get(i);
			Map documentInfoMap = allDocumentInfoList.get(i);
			DomainObject route = routeList.get(i);
			Map printSettingMap = allPrintSettingList.get(i);
			MapList printSetParamsList = allPrintSetParamsList.get(i);
			printDocument(document, documentInfoMap, route, printSettingMap, printSetParamsList);
		}
	}

	/**
	 * signature current document
	 * 
	 * @param context
	 *            The Matrix Context.
	 * @param DomainObject
	 *            context object
	 * @param String
	 *            related Route objectId
	 * @param String
	 *            file path
	 * @param String
	 *            temp file path
	 * @param StringList
	 *            baseline objectId
	 * @return void
	 * @throws Exception
	 *             If the operation fails.
	 */
	private void printDocument(DomainObject document, Map documentInfoMap, DomainObject route, Map printSettingMap,
			MapList printSetParamsList) throws Exception {
		String relPattern = DomainConstants.RELATIONSHIP_ROUTE_NODE;
		String typePattern = DomainConstants.TYPE_PERSON;

		StringList busSelects = new StringList(DomainConstants.SELECT_ID);
		busSelects.addElement(DomainConstants.SELECT_NAME);
		StringList relSelects = new StringList(DomainConstants.SELECT_RELATIONSHIP_ID);
		relSelects.addElement(AttributeConstants.Select_Route_Sequence);
		relSelects.addElement(AttributeConstants.Select_Actual_Completion_Date);
		relSelects.addElement(AttributeConstants.Select_Approval_Status);
		relSelects.addElement(AttributeConstants.Select_Title);

		MapList mpRouteTask = new MapList();
		if (route != null) {
			mpRouteTask = route.getRelatedObjects(context, relPattern, typePattern, busSelects, relSelects, false,
					true, (short) 1, null, null, (short) 0);
			mpRouteTask.sort(AttributeConstants.Select_Actual_Completion_Date, "ascending", "String");
		}

		System.out.println("printDocument:mpRouteTask:" + mpRouteTask);
		StringList phaseIdentifyList = document.getInfoList(context, "to["
				+ RelationShipConstants.GW_PhaseBaseLineContent + "].from.attribute["
				+ AttributeConstants.GW_PhaseIdentify + "].value");

		FileList files = document.getFiles(context);
		if (Util.isNotEmptyList(files)) {
			this.attachmentCount = this.attachmentCount + files.size();
			for (int i = 0; i < files.size(); i++) {
				String strFileName = files.getElement(i).getName();
				String strFileFormat = files.getElement(i).getFormat();
				System.out.println("strFileName++++++" + strFileName);
				if (strFileName.toLowerCase().contains(".pdf")) {
					document.checkoutFile(context, false, strFileFormat, strFileName, app_path + temp_path);
					printPDFDocument(documentInfoMap, strFileName, mpRouteTask, route, phaseIdentifyList,
							printSettingMap, printSetParamsList);
				} else if (strFileName.toLowerCase().contains(".doc") || strFileName.toLowerCase().contains(".docx")) {
					document.checkoutFile(context, false, strFileFormat, strFileName, app_path + temp_path);
					strFileName = transferWORD2PDF(strFileName);
					printPDFDocument(documentInfoMap, strFileName, mpRouteTask, route, phaseIdentifyList,
							printSettingMap, printSetParamsList);
				} else if (strFileName.toLowerCase().contains(".xls") || strFileName.toLowerCase().contains(".xlsx")) {
					document.checkoutFile(context, false, strFileFormat, strFileName, app_path + temp_path);
					strFileName = transferExcel2PDF(strFileName);
					printPDFDocument(documentInfoMap, strFileName, mpRouteTask, route, phaseIdentifyList,
							printSettingMap, printSetParamsList);
				}

			}
		}
	}

	/**
	 * transfer office file to PDF format
	 * 
	 * @param context
	 *            The Matrix Context.
	 * @param String
	 *            file path.
	 * @param String
	 *            file name.
	 * @return String
	 * @throws Exception
	 *             If the operation fails.
	 */

	private String transferWORD2PDF(String strFileName) throws Exception {
		// 判断转化工具是否在使用
		if (isConvertUsing) {
			Thread.sleep(5000);
		} else {
			isConvertUsing = true;
		}
		System.out.println("Start Word...");
		long start = System.currentTimeMillis();
		ActiveXComponent app = null;
		Dispatch doc = null;
		String strDestinationName = "";
		try {
			app = new ActiveXComponent("Word.Application");
			app.setProperty("Visible", new Variant(false));

			Dispatch docs = app.getProperty("Documents").toDispatch();

			String sfileName = app_path + temp_path + "\\" + strFileName;

			strDestinationName = strFileName.replace(".docx", ".pdf").replace(".doc", ".pdf");

			String toFileName = app_path + temp_path + "\\" + strDestinationName;

			doc = Dispatch.call(docs, "Open", sfileName).toDispatch();
			System.out.println("open File..." + sfileName);
			System.out.println("Transfer to PDF..." + toFileName);
			File tofile = new File(toFileName);
			if (tofile.exists()) {
				tofile.delete();
			}
			Dispatch.call(doc, "SaveAs", toFileName, wdFormatPDF);
			long end = System.currentTimeMillis();
			System.out.println("end..spent" + (end - start) + "ms.");

		} catch (Exception e) {
			e.printStackTrace();
			throw new Exception("word\u8F6CPDF\u51FA\u9519");
		} finally {
			Dispatch.call(doc, "Close", false);
			System.out.println("close file");
			if (app != null) {
				app.invoke("Quit", new Variant[] {});
			}
			isConvertUsing = false;
		}
		ComThread.Release();
		return strDestinationName;

	}

	/**
	 * transfer office file to PDF format
	 * 
	 * @param context
	 *            The Matrix Context.
	 * @param String
	 *            file path.
	 * @param String
	 *            file name.
	 * @return String
	 * @throws Exception
	 *             If the operation fails.
	 */

	private String transferExcel2PDF(String strFileName) throws Exception {

		// 判断转化工具是否在使用
		if (isConvertUsing) {
			Thread.sleep(5000);
		} else {
			isConvertUsing = true;
		}
		System.out.println("Start Excel...");
		long start = System.currentTimeMillis();
		ActiveXComponent app = null;
		Dispatch workbook = null;
		String strDestinationName = "";
		try {
			app = new ActiveXComponent("Excel.Application");
			app.setProperty("Visible", new Variant(false));

			Dispatch workbooks = app.getProperty("Workbooks").toDispatch();

			String sfileName = app_path + temp_path + "\\" + strFileName;

			strDestinationName = strFileName.replace(".xlsx", ".pdf").replace(".xls", ".pdf");

			String toFileName = app_path + temp_path + "\\" + strDestinationName;

			workbook = Dispatch.invoke(workbooks, "Open", Dispatch.Method,
					new Object[] { sfileName, new Variant(false), new Variant(false) }, new int[3]).toDispatch();
			
			Dispatch sheets = Dispatch.get(workbook, "sheets").toDispatch(); 
			setPrintArea(sheets);
			System.out.println("open File..." + sfileName);
			System.out.println("Transfer to PDF..." + toFileName);
			File tofile = new File(toFileName);
			if (tofile.exists()) {
				tofile.delete();
			}
			// 0－PDF 1－xps
			Dispatch.call(workbook, "ExportAsFixedFormat", 0, toFileName);
			long end = System.currentTimeMillis();
			System.out.println("end..spent" + (end - start) + "ms.");

		} catch (Exception e) {
			e.printStackTrace();
			throw new Exception("word\u8F6CPDF\u51FA\u9519");
		} finally {
			Dispatch.call(workbook, "Close", false);
			System.out.println("close file");
			if (app != null) {
				app.invoke("Quit", new Variant[] {});
			}
			isConvertUsing = false;
		}
		ComThread.Release();
		return strDestinationName;

	}

	   /* 
	    *  为每个表设置打印区域 
	    */  
	   private void setPrintArea(Dispatch sheets){  
	       int count = Dispatch.get(sheets, "count").changeType(Variant.VariantInt).getInt();  
	       for (int i = count; i >= 1; i--) {  
	              Dispatch sheet = Dispatch.invoke(sheets, "Item",  
	                      Dispatch.Get, new Object[] { i }, new int[1]).toDispatch();  
	          Dispatch page = Dispatch.get(sheet, "PageSetup").toDispatch();  
	          Dispatch.put(page, "PrintArea", false);  
	          Dispatch.put(page, "Orientation", 1);  //横纵打印
	          Dispatch.put(page, "Zoom", false);      //值为100或false  
	          Dispatch.put(page, "FitToPagesTall", 1);  //所有行为一页  
	          Dispatch.put(page, "FitToPagesWide", 1);     //所有列为一页(1或false)
	       }
	   }   
	
	/**
	 * insert content to PDF
	 * 
	 * @param context
	 *            The Matrix Context.
	 * @param String
	 *            app server path.
	 * @param String
	 *            temporary location.
	 * @param String
	 *            file name.
	 * @param MapList
	 *            route task info
	 * @param String
	 *            revision.
	 * @param StringList
	 *            phaseIDList
	 * @return Map
	 * @throws Exception
	 *             If the operation fails.
	 */

	private void printPDFDocument(Map documentInfoMap, String strFileName, MapList mpRouteTask, DomainObject domRoute,
			StringList phaseIdentifyList, Map printSettingMap, MapList printSetParamsList) throws Exception {
		PdfReader reader = null;
		PdfStamper stamper = null;
		try {
			String strPageNumber = (String) printSettingMap.get(AttributeConstants.Select_GW_PrintPageNumber);
			System.out.println("printPDFDocument::strPageNumber:" + strPageNumber);
			reader = new PdfReader(app_path + temp_path + "\\" + strFileName, "PDF".getBytes());
			// create
			String format = (String) documentInfoMap.get(AttributeConstants.Select_GW_Format);
			if (UIUtil.isNotNullAndNotEmpty(format)) {
				FileUtil.createFolder(app_path + temp_path + "\\output\\" + format);
				stamper = new PdfStamper(reader, new FileOutputStream(app_path + temp_path + "\\output\\" + format
						+ "\\" + strFileName));
			} else {
				FileUtil.createFolder(app_path + temp_path + "\\output");
				stamper = new PdfStamper(reader,
						new FileOutputStream(app_path + temp_path + "\\output\\" + strFileName));
			}

			Set<Integer> pageNumberSet = new HashSet<Integer>();
			int totalPageNumber = reader.getNumberOfPages();
			System.out.println("----------------------------------------------------------------------->totalPageNumber:"+totalPageNumber);
			boolean isSingleConfNumer = false;

			if (UIUtil.isNotNullAndNotEmpty(strPageNumber)) {
				pageNumberSet = getPageNumberSet(strPageNumber, totalPageNumber);
				System.out.println("=========================================================================>pageNumberSet:"+pageNumberSet);
			} else {
				isSingleConfNumer = true;
			}
			System.out.println("printPDFDocument::pageNumberSet:" + pageNumberSet);
			System.out.println("printPDFDocument::isSingleConfNumer:" + isSingleConfNumer);
			System.out.println("printPDFDocument::phaseIdentifyList:" + phaseIdentifyList);
			printPDFDocument(documentInfoMap, stamper, pageNumberSet, printSetParamsList, mpRouteTask, domRoute,
					phaseIdentifyList, isSingleConfNumer, totalPageNumber);
		} catch (Exception e) {
			e.printStackTrace();
			throw new FrameworkException(e);
		} finally {
			if (stamper != null) {
				stamper.close();
			}
			if (reader != null) {
				reader.close();
			}
		}
	}

	private void printPDFDocument(Map documentInfoMap, PdfStamper stamper, Set<Integer> pageNumberSet,
			MapList printSetParamsList, MapList mpRouteTask, DomainObject domRoute, StringList phaseIdentifyList,
			boolean isSingleConfNumer, int totalPageNumber) throws Exception {
		System.out.println("printPDFDocument:pageNumberSet:" + pageNumberSet);
		List<Map<String, Object>> insertInfoMap = new ArrayList<Map<String, Object>>();
		MapList hasTwoReviewParams = new MapList();
		for (int i = 0; i < printSetParamsList.size(); i++) {
			Map printSetParamsMap = (Map) printSetParamsList.get(i);
			String printParamTag = (String) printSetParamsMap.get(AttributeConstants.Select_GW_PrintTag);
			if (printParamTag.equalsIgnoreCase("revision")) {
				String strRevision = (String) documentInfoMap.get(DomainConstants.SELECT_REVISION);
				getPrintInfoForText(printSetParamsMap, insertInfoMap, pageNumberSet, isSingleConfNumer,
						totalPageNumber, strRevision, 10);
			} else if (printParamTag.startsWith("phase")) {
				String phase = printParamTag.substring("phase".length()).toUpperCase();
				if (phaseIdentifyList.contains(phase)) {
					getPrintInfoForText(printSetParamsMap, insertInfoMap, pageNumberSet, isSingleConfNumer,
							totalPageNumber, phase, 10);
				}
			} else if (printParamTag.equalsIgnoreCase("completiondate0")) {//流程发起时间
				if (!OtherConstants.NOSIGN.equals(this.stamperMap.get("printIsSign")) && domRoute != null) {
					String strCreationDate = sdf.format(new Date((String) domRoute.getInfo(context,
							DomainConstants.SELECT_ORIGINATED)));
					getPrintInfoForText(printSetParamsMap, insertInfoMap, pageNumberSet, isSingleConfNumer,
							totalPageNumber, strCreationDate, 6);
				}
			} else if (printParamTag.equalsIgnoreCase("routenode0")) {//流程发起人
				if (!OtherConstants.NOSIGN.equals(this.stamperMap.get("printIsSign")) && domRoute != null) {
					String strImageName = domRoute.getInfo(context, DomainConstants.SELECT_ORIGINATOR);
					getPrintInfoForImage(printSetParamsMap, insertInfoMap, pageNumberSet, isSingleConfNumer,
							totalPageNumber, strImageName);
				}
			} else if (printParamTag.startsWith("completiondate")) {
				if (!OtherConstants.NOSIGN.equals(this.stamperMap.get("printIsSign"))) {
					String printSeq = (String) printSetParamsMap.get(AttributeConstants.Select_GW_PrintSequence);
					String[] printSeqArray = printSeq.split("\\|");
					if (printSeqArray.length > 1) {
						hasTwoReviewParams.add(printSetParamsMap);
					} else {
						List<String> printSeqList = new ArrayList<String>();
						printSeqList.add(printSeq.trim());
						getPrintInfoForReview(insertInfoMap, mpRouteTask, pageNumberSet, printSetParamsList,
								printSetParamsMap, printSeqList, isSingleConfNumer, totalPageNumber);
					}
				}
			} else if (printParamTag.equalsIgnoreCase("stamper")
					&& StringUtils.isNotEmpty(stamperMap.get("printApplication"))
					&& StringUtils.isNotEmpty(stamperMap.get("printDepartment"))) {
				String format = (String) documentInfoMap.get(AttributeConstants.Select_GW_Format);
				getPrintInfoForStamperText(printSetParamsMap, insertInfoMap, pageNumberSet, isSingleConfNumer,
						totalPageNumber, 10, format);
			} else if (printParamTag.equalsIgnoreCase("water")) {
				getPrintInfoForWaterText(printSetParamsMap, insertInfoMap, pageNumberSet, isSingleConfNumer,
						totalPageNumber, 10);
			}
		}
		// 如果打印参数中的‘打印序列’中配置了多个流程节点，单独处理
		System.out.println("printPDFDocument:hasTwoReviewParams:" + hasTwoReviewParams);
		if (!OtherConstants.NOSIGN.equals(this.stamperMap.get("printIsSign"))) {
			for (int i = 0; i < hasTwoReviewParams.size(); i++) {
				Map printSetParamsMap = (Map) hasTwoReviewParams.get(i);
				String printParamTag = (String) printSetParamsMap.get(AttributeConstants.Select_GW_PrintTag);
				if (printParamTag.startsWith("completiondate")) {
					String printSeq = (String) printSetParamsMap.get(AttributeConstants.Select_GW_PrintSequence);
					String[] printSeqArray = printSeq.split("\\|");
					List<String> printSeqList = new ArrayList<String>();
					for (String str : printSeqArray) {
						printSeqList.add(str.trim());
					}
					getPrintInfoForReview(insertInfoMap, mpRouteTask, pageNumberSet, printSetParamsList,
							printSetParamsMap, printSeqList, isSingleConfNumer, totalPageNumber);
				}
			}
		}

		System.out.println("insertInfoMap:**************************************:" + insertInfoMap);
		System.out.println("pageNumberSet:**************************************:" + pageNumberSet);
		insertPDFContextValue(stamper, insertInfoMap, pageNumberSet);
		// 记录
		if ("Archives".equals((String) this.stamperMap.get("printAction"))) {
			createPrintRecord();
		}
	}

	/**
	 * 获取需要写入签审的信息
	 */
	private void getPrintInfoForReview(List<Map<String, Object>> insertInfoMap, MapList mpRouteTask,
			Set<Integer> pageNumberSet, MapList printSetParamsList, Map printSetParamsMap, List<String> printSeqList,
			boolean isSingleConfNumer, int totalPageNumber) throws Exception {
		System.out.println("getPrintInfoForReview:printSeqList:" + printSeqList);
		String printParamTag = (String) printSetParamsMap.get(AttributeConstants.Select_GW_PrintTag);
		for (int k = 0; k < mpRouteTask.size(); k++) {
			Map routeTaskMap = (Map) mpRouteTask.get(k);
			String title = (String) routeTaskMap.get(AttributeConstants.Select_Title);
			System.out.println("===============================>title:"+title);
			System.out.println("===============================>printSeqList:"+printSeqList.toString());
			if (!printSeqList.contains(title)) {
				continue;
			}

			String strCompletionDate = (String) routeTaskMap.get(AttributeConstants.Select_Actual_Completion_Date);
			String strApprovalStatus = (String) routeTaskMap.get(AttributeConstants.Select_Approval_Status);
			String strImageName = (String) routeTaskMap.get(DomainConstants.SELECT_NAME);
			if (!"Approve".equalsIgnoreCase(strApprovalStatus)) {
				continue;
			}
			if (strCompletionDate == null || "".equals(strCompletionDate)) {
				continue;
			}
			strCompletionDate = sdf.format(new Date(strCompletionDate));
			getPrintInfoForText(printSetParamsMap, insertInfoMap, pageNumberSet, isSingleConfNumer, totalPageNumber,
					strCompletionDate, 6);

			Map routeNodeParamsMap = null;
			// 遍历获取到签审对应的审核人
			for (int j = 0; j < printSetParamsList.size(); j++) {
				Map printSetParamsMap1 = (Map) printSetParamsList.get(j);
				String printParamTag1 = (String) printSetParamsMap1.get(AttributeConstants.Select_GW_PrintTag);
				if (printParamTag1.equalsIgnoreCase("routenode" + printParamTag.substring("completiondate".length()))) {
					routeNodeParamsMap = printSetParamsMap1;
					System.out.println(printParamTag1 + "   " + printParamTag + "   " + strImageName);
					break;
				}
			}
			if (routeNodeParamsMap != null) {
				getPrintInfoForImage(routeNodeParamsMap, insertInfoMap, pageNumberSet, isSingleConfNumer,
						totalPageNumber, strImageName);
			}
			// mpRouteTask.remove(routeTaskMap);
			break;
		}
	}

	/**
	 * 获取需要写入图片的信息
	 */
	private void getPrintInfoForImage(Map positionMap, List<Map<String, Object>> insertInfoMap,
			Set<Integer> pageNumberSet, boolean isSingleConfNumer, int totalPageNumber, String strImageName)
			throws Exception {
		int strXcoord = Integer.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintPosition_x));
		int strYcoord = Integer.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintPosition_y));
		int strImgWidth = Integer.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintImgWidth));
		int strImgHeight = Integer.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintImgHeight));
		int strFontSize = Integer.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintFontSize));
		// boolean isImage = Boolean.parseBoolean((String) positionMap.get(AttributeConstants.Select_GW_PrintIsImg));
		String signMethod = this.stamperMap.get("printIsSign");
		Map<String, Object> map = new HashMap<String, Object>();
		if (OtherConstants.SIGNBYHAND.equalsIgnoreCase(signMethod)) {
			String imagePath = Util.getTemplateFilePath("template\\signature\\" + strImageName + ".jpg");
			Image image = Image.getInstance(imagePath);
			image.setAbsolutePosition(strXcoord, strYcoord);

			strImgWidth = strImgWidth == 0 ? 150 : strImgWidth;
			strImgHeight = strImgHeight == 0 ? 10 : strImgHeight;
			image.scaleToFit(strImgWidth, strImgHeight);
			map.put("Image", image);
		} else {
			String personId = "";
			MapList personList = PersonUtil.getPersonByName(context, strImageName);
			if (Util.isNotEmptyList(personList)) {
				personId = (String) ((Map) personList.get(0)).get(DomainConstants.SELECT_ID);
			}
			DomainObject person = DomainObject.newInstance(context, personId);
			String firstName = person.getInfo(context, AttributeConstants.Select_First_Name);
			strFontSize = strFontSize == 0 ? 10 : strFontSize;
			map.put("FontSize", strFontSize);
			map.put("Xcoord", strXcoord);
			map.put("Ycoord", strYcoord);
			map.put("Text", firstName);
		}
		if (isSingleConfNumer) {
			String printPageNumber = (String) positionMap.get(AttributeConstants.Select_GW_PrintPageNumber);
			Set<Integer> paramPageNumSet = getPageNumberSet(printPageNumber, totalPageNumber);
			pageNumberSet.addAll(paramPageNumSet);
			map.put("PageNumberSet", paramPageNumSet);
		} else {
			map.put("PageNumberSet", pageNumberSet);
		}

		insertInfoMap.add(map);
	}

	/**
	 * 获取需要写入字符串的信息
	 */
	private void getPrintInfoForText(Map positionMap, List<Map<String, Object>> insertInfoMap,
			Set<Integer> pageNumberSet, boolean isSingleConfNumer, int totalPageNumber, String textValue,
			int defaultFontSize) throws Exception {

		int strXcoord = Integer.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintPosition_x));
		int strYcoord = Integer.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintPosition_y));
		int strFontSize = Integer.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintFontSize));

		strFontSize = strFontSize == 0 ? defaultFontSize : strFontSize;

		Map<String, Object> map = new HashMap<String, Object>();
		map.put("FontSize", strFontSize);
		map.put("Xcoord", strXcoord);
		map.put("Ycoord", strYcoord);
		map.put("Text", textValue);
		if (isSingleConfNumer) {
			String printPageNumber = (String) positionMap.get(AttributeConstants.Select_GW_PrintPageNumber);
			Set<Integer> paramPageNumSet = getPageNumberSet(printPageNumber, totalPageNumber);
			pageNumberSet.addAll(paramPageNumSet);
			map.put("PageNumberSet", paramPageNumSet);
		} else {
			map.put("PageNumberSet", pageNumberSet);
		}
		insertInfoMap.add(map);
	}

	/**
	 * 获取需要写入图章的信息
	 */
	private void getPrintInfoForStamperText(Map positionMap, List<Map<String, Object>> insertInfoMap,
			Set<Integer> pageNumberSet, boolean isSingleConfNumer, int totalPageNumber, int defaultFontSize,
			String format) throws Exception {

		System.out.println("--------------------------mdsmdsmds-------------------------------------");
		int strXcoord = Integer.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintPosition_x));
		int strYcoord = Integer.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintPosition_y));
		if (strXcoord == 0 && strYcoord == 0) {
			switch (format) {
			case "A4":
				strXcoord = 38;
				strYcoord = 250;
				break;
			case "A3":
				strXcoord = 38;
				strYcoord = 250;
				break;
			case "A2":
				strXcoord = 38;
				strYcoord = 250;
				break;
			case "A1":
				strXcoord = 38;
				strYcoord = 250;
				break;
			}
		}
		System.out.println("--------------------------mdsmdsmds-------------------------------------strXcoord:"
				+ strXcoord + " strYcoord:" + strYcoord);
		int strStamperWidth = Integer
				.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintStamperWidth));

		Map<String, Object> map = new HashMap<String, Object>();
		map.put("StamperWidth", strStamperWidth);
		map.put("Xcoord", strXcoord);
		map.put("Ycoord", strYcoord);
		map.put("ApplicationText", stamperMap.get("printApplication"));
		map.put("DateText", getDate(stamperMap.get("printDate")));
		map.put("DepartmentText", stamperMap.get("printDepartment"));
		map.put("Color", "Black");
		if (isSingleConfNumer) {
			String printPageNumber = (String) positionMap.get(AttributeConstants.Select_GW_PrintPageNumber);
			Set<Integer> paramPageNumSet = getPageNumberSet(printPageNumber, totalPageNumber);
			pageNumberSet.addAll(paramPageNumSet);
			map.put("PageNumberSet", paramPageNumSet);
		} else {
			map.put("PageNumberSet", pageNumberSet);
		}
		insertInfoMap.add(map);
	}

	/**
	 * 获取需要写入水印的信息
	 */
	private void getPrintInfoForWaterText(Map positionMap, List<Map<String, Object>> insertInfoMap,
			Set<Integer> pageNumberSet, boolean isSingleConfNumer, int totalPageNumber, int defaultFontSize)
			throws Exception {

		int strXcoord = Integer.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintPosition_x));
		int strYcoord = Integer.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintPosition_y));
		int strFontSize = Integer.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintFontSize));
		int strRotation = Integer.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintRotation));
		int strOpacity = Integer.parseInt((String) positionMap.get(AttributeConstants.Select_GW_PrintOpacity));

		strFontSize = strFontSize == 0 ? defaultFontSize : strFontSize;

		String textValue = stamperMap.get("printWater");

		Map<String, Object> map = new HashMap<String, Object>();
		map.put("FontSize", strFontSize);
		map.put("Xcoord", strXcoord);
		map.put("Ycoord", strYcoord);
		map.put("WaterText", textValue);
		map.put("Rotation", strRotation);
		map.put("Opacity", strOpacity);
		if (isSingleConfNumer) {
			String printPageNumber = (String) positionMap.get(AttributeConstants.Select_GW_PrintPageNumber);
			Set<Integer> paramPageNumSet = getPageNumberSet(printPageNumber, totalPageNumber);
			pageNumberSet.addAll(paramPageNumSet);
			map.put("PageNumberSet", paramPageNumSet);
		} else {
			map.put("PageNumberSet", pageNumberSet);
		}
		insertInfoMap.add(map);
	}

	/**
	 * 把需要写入的信息写入PDF
	 */
	private void insertPDFContextValue(PdfStamper stamper, List<Map<String, Object>> insertInfoMap,
			Set<Integer> pageNumberSet) throws Exception {
		Font font = new Font(BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED), 8, Font.BOLD);
		if (insertInfoMap.size() != 0) {
			// 循环总共需要写入信息的页码，把对应的信息写入该页
			for (int pageNumber : pageNumberSet) {
				PdfContentByte over = stamper.getOverContent(pageNumber);
				if (over == null) {
					continue;
				}
				for (Map<String, Object> inertInfo : insertInfoMap) {
					System.out.println("=======================================>inertInfo:"+inertInfo);
					if (inertInfo.size() == 2) {
						Set<Integer> paramNumberSet = (Set<Integer>) inertInfo.get("PageNumberSet");
						if (paramNumberSet.contains(pageNumber)) {
							over.addImage((Image) inertInfo.get("Image"));
						}
					} else if (inertInfo.size() == 5) {
						Set<Integer> paramNumberSet = (Set<Integer>) inertInfo.get("PageNumberSet");
						if (paramNumberSet.contains(pageNumber)) {
							// 内容
							over.beginText();
							over.setFontAndSize(font.getBaseFont(), (Integer) inertInfo.get("FontSize"));
							over.setColorFill(BaseColor.BLACK);
							over.setTextMatrix((Integer) inertInfo.get("Xcoord"), (Integer) inertInfo.get("Ycoord"));
							over.showText((String) inertInfo.get("Text"));
							System.out.println(">>>>>>>>>>>>>>>写：" + inertInfo.get("Text"));
							over.endText();
						}
					} else if (inertInfo.size() == 8) {
						Set<Integer> paramNumberSet = (Set<Integer>) inertInfo.get("PageNumberSet");
						if (paramNumberSet.contains(pageNumber)) {
							// 图章
							over.beginText();
							over.setColorFill(BaseColor.BLACK);
							PdfPTable table = new PdfPTable(1);
							float[] width = { (Integer) inertInfo.get("StamperWidth") };
							table.setTotalWidth(width);
							table.addCell(new PdfPCell(new Paragraph((String) inertInfo.get("ApplicationText"), font)));
							table.addCell(new PdfPCell(new Paragraph((String) inertInfo.get("DateText"), font)));
							table.addCell(new PdfPCell(new Paragraph((String) inertInfo.get("DepartmentText"), font)));
							table.writeSelectedRows(0, -1, (Integer) inertInfo.get("Xcoord"),
									(Integer) inertInfo.get("Ycoord"), over);
							over.endText();
						}

					} else if (inertInfo.size() == 7) {
						Set<Integer> paramNumberSet = (Set<Integer>) inertInfo.get("PageNumberSet");
						if (paramNumberSet.contains(pageNumber)) {
							// 水印
							PdfGState gs = new PdfGState();
							over.beginText();
							gs.setFillOpacity(((Integer) inertInfo.get("Opacity")) / 100);
							over.setColorFill(BaseColor.LIGHT_GRAY);
							over.setFontAndSize(font.getBaseFont(), (Integer) inertInfo.get("FontSize"));
							over.showTextAligned(Element.ALIGN_CENTER, (String) inertInfo.get("WaterText"),
									(Integer) inertInfo.get("Xcoord"), (Integer) inertInfo.get("Ycoord"),
									(Integer) inertInfo.get("Rotation"));
							over.endText();
						}

					}
				}
			}
		}
	}

	/**
	 * 获取需要写入信息的页码 ,页码: |代表分割 *代表全部 -代表区域
	 */
	private Set<Integer> getPageNumberSet(String strPage, int totalPage) throws Exception {
		Set<Integer> pageNumberSet = new HashSet<Integer>();
		String[] strArray = strPage.split("\\|");

		for (String childStr : strArray) {
			Pattern pattern = Pattern.compile("^[0-9]+\\-([0-9]+|\\*+)$");
			Pattern pattern1 = Pattern.compile("^[0-9]+$");
			if (childStr.equals("*")) {
				for (int i = 1; i <= totalPage; i++) {
					pageNumberSet.add(i);
				}
				return pageNumberSet;
			} else if (pattern.matcher(childStr).matches()) {
				String childStrArray[] = childStr.split("-");
				int firstNumber = Integer.valueOf(childStrArray[0]);
				int lastNumber = 0;
				if (firstNumber <= totalPage) {
					if (childStrArray[1].equals("*")) {
						lastNumber = totalPage;
						for (int i = firstNumber; i <= lastNumber; i++) {
							pageNumberSet.add(i);
						}
					} else {
						lastNumber = Integer.valueOf(childStrArray[1]);
						if (lastNumber >= firstNumber) {
							if (lastNumber >= totalPage) {
								lastNumber = totalPage;
							}
							for (int i = firstNumber; i <= lastNumber; i++) {
								pageNumberSet.add(i);
							}
						}
					}
				}
			} else if (pattern1.matcher(childStr).matches()) {
				int number = Integer.valueOf(childStr);
				if (number <= totalPage) {
					pageNumberSet.add(number);
				}
			}
		}
		return pageNumberSet;
	}

	private void getErrorInfo(Context context, Map<String, String> returnMap, String errorName, String errorKey) {
		try {
			String returnMessage = EnoviaResourceBundle.getProperty(context, "EngineeringCentral", errorKey, context
					.getSession().getLanguage());
			returnMap.put(OtherConstants.ReturnResult, OtherConstants.FAIL);
			returnMap.put(OtherConstants.ReturnMessage, errorName + "\\n" + returnMessage);
		} catch (FrameworkException e) {
			e.printStackTrace();
		}
	}

	public void createPrintRecord() throws Exception {
		try {
			ContextUtil.startTransaction(context, true);
			DomainObject printRecord = new DomainObject();
			printRecord.createObject(context, TypeConstants.GW_PrintRecord, null, null, PolicyConstants.GW_PrintRecord,
					"");
			printRecord.setAttributeValue(context, AttributeConstants.GW_PrintDate,
					getDate(stamperMap.get("printDate")));
			printRecord.setAttributeValue(context, AttributeConstants.GW_PrintApplication,
					stamperMap.get("printApplication"));
			printRecord.setAttributeValue(context, AttributeConstants.GW_PrintDepartment,
					stamperMap.get("printDepartment"));
			printRecord.setAttributeValue(context, AttributeConstants.GW_PrintAttachmentCount, this.attachmentCount
					+ "");
			printRecord.setAttributeValue(context, AttributeConstants.GW_PrintApplicant,
					stamperMap.get("printApplicant"));
			printRecord.setAttributeValue(context, AttributeConstants.GW_PrintUser,
					stamperMap.get("printUser"));
			ContextUtil.commitTransaction(context);
		} catch (Exception e) {
			ContextUtil.abortTransaction(context);
			e.printStackTrace();
		}
	}

	private String getDate(String dateString) {
		if (StringUtils.isNotBlank(dateString)) {
			return dateString;
		} else {
			return new SimpleDateFormat(OtherConstants.DATE_FORMAT).format(new Date());
		}
	}

}

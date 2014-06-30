package es.ramondin.compdec.exportarexcel.view.util;

import es.ramondin.util.general.RmdColor;

import java.awt.Color;

import java.lang.reflect.Field;

import java.text.ParseException;
import java.text.SimpleDateFormat;

import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import java.util.Map;

import javax.el.ELContext;
import javax.el.ExpressionFactory;
import javax.el.ValueExpression;

import javax.faces.component.UIComponent;
import javax.faces.context.FacesContext;

import oracle.adf.view.rich.component.rich.data.RichColumn;
import oracle.adf.view.rich.component.rich.data.RichTable;

import oracle.jbo.uicli.binding.JUCtrlHierNodeBinding;

import org.apache.commons.lang.StringUtils;
import org.apache.myfaces.trinidad.component.UIXValue;
import org.apache.myfaces.trinidad.model.CollectionModel;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExportarExcelUtil {
    public ExportarExcelUtil() {
        super();
    }

    public static HashMap<String, String> getMapaEstilos(String inlineStyle) {
        HashMap<String, String> res = null;

        if (inlineStyle != null) {
            String filasStyle[] = inlineStyle.split(";");
            res = new HashMap<String, String>();
            String attrStyle[];

            for (int k = 0; k < filasStyle.length; k++) {
                attrStyle = filasStyle[k].split(":");

                if (attrStyle.length == 2)
                    res.put(attrStyle[0], attrStyle[1]);
            }
        }

        return res;
    }

    public static Color getColor(HashMap<String, String> mapaEstilos, String clave) {
        Color res = null;

        if (mapaEstilos != null) {
            String txtColor = mapaEstilos.get(clave);

            if (txtColor != null) {
                Field field;
                try {
                    field = Class.forName("java.awt.Color").getField(txtColor.toLowerCase());
                    res = (Color)field.get(null);
                } catch (ClassNotFoundException e) {
                    e = null;
                } catch (IllegalAccessException e) {
                    e = null;
                } catch (NoSuchFieldException e) {
                    e = null;
                }

                if (res == null) {
                    try {
                        res = RmdColor.decode(txtColor);
                    } catch (NumberFormatException e) {
                        e = null;
                    }
                }
            }
        }

        return res;
    }
    
    private static void aplicaColor(XSSFCell celda, XSSFCellStyle celdaStyle, Color color) {
        celdaStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        celdaStyle.setFillForegroundColor(new XSSFColor(color));
        celdaStyle.setFillBackgroundColor(new XSSFColor(color));
        celda.setCellStyle(celdaStyle);
    }

    public static XSSFWorkbook generarExcel(FacesContext facesContext, RichTable tabla, String nombreHoja, Boolean mostrarColsOcultas, ArrayList<Integer> columnasFecRMD) {
        XSSFWorkbook libro = new XSSFWorkbook();

        XSSFSheet hoja = libro.createSheet(WorkbookUtil.createSafeSheetName(nombreHoja));
        XSSFRow fila;
        XSSFCell celda;
        XSSFCellStyle celdaStyle;
        XSSFFont defaultFont = libro.createFont();
        XSSFFont defaultFontBold = libro.createFont();
        defaultFontBold.setBold(true);

        List<RichColumn> cols = new ArrayList<RichColumn>();
        RichColumn col;
        for (UIComponent c : tabla.getChildren()) {
            if (c instanceof RichColumn) {
                col = (RichColumn)c;
                if (mostrarColsOcultas || (col.isRendered() && col.isVisible()))
                    cols.add(col);
            }
        }

        //Cabecera
        fila = hoja.createRow(0);
        for (int i = 0; i < cols.size(); i++) {
            col = cols.get(i);

            celda = fila.createCell(i);
            celdaStyle = libro.createCellStyle();
            
            celdaStyle.setFont(defaultFontBold);
            celdaStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
            
            celda.setCellStyle(celdaStyle);
            celda.setCellValue(StringUtils.defaultString(col.getHeaderText()));
        }
        hoja.createFreezePane(0, 1, 0, 1);

        ELContext elContext = facesContext.getELContext();
        ExpressionFactory expressionFactory = facesContext.getApplication().getExpressionFactory();
        /*        try {
            if (StringUtils.isNotEmpty(tabla.getVarStatus())) {
                // create varStatusMap (this is the only way I have found, if you have another solution be my guest :))
                Method m;
                m = UIXIterator.class.getDeclaredMethod("createVarStatusMap", new Class[0]);

                m.setAccessible(true);
                Object varStatus;
                // inyect 'varStatus' into ELContext
                varStatus = m.invoke(tabla, new Object[0]);
                String el = String.format("#{%s}", tabla.getVarStatus());
                ValueExpression exp = expressionFactory.createValueExpression(elContext, el, Object.class);
                exp.setValue(elContext, varStatus);
            }
        } catch (NoSuchMethodException e) {
        } catch (IllegalAccessException e) {
        } catch (InvocationTargetException e) {
        }
        */

        //Filas
        CollectionModel model = (CollectionModel)tabla.getValue();
        int rowcount = model.getRowCount();
        boolean esFechaRMD = false;
        SimpleDateFormat sdfFormato = new SimpleDateFormat("yyyyMMdd");
        String fechaRMDTxt = null;
        Date fechaRMD = null;
               
        for (int i = 0; i < rowcount; i++) {
            model.setRowIndex(i);

            JUCtrlHierNodeBinding row = (JUCtrlHierNodeBinding)model.getRowData();
            if (StringUtils.isNotEmpty(tabla.getVar())) {
                String el = String.format("#{%s}", tabla.getVar());
                ValueExpression exp = expressionFactory.createValueExpression(elContext, el, Object.class);
                exp.setValue(elContext, row);
            }

            fila = hoja.createRow(i + 1);
            for (int j = 0; j < cols.size(); j++) {
                col = cols.get(j);
                celda = fila.createCell(j);
                celdaStyle = libro.createCellStyle();

                ValueExpression inlineStyleVE = col.getValueExpression("inlineStyle");
                ValueExpression styleClassVE = col.getValueExpression("styleClass");
                ValueExpression alignVE = col.getValueExpression("align");
                String style = inlineStyleVE == null ? "" : (String)inlineStyleVE.getValue(facesContext.getELContext());
                String styleClass = styleClassVE == null ? "" : (String)styleClassVE.getValue(facesContext.getELContext());
                String align = alignVE == null ? "" : (String)alignVE.getValue(facesContext.getELContext());

                //Tratamos los estilos
                Color colorEstilo = null;
                
                if (styleClass.contentEquals("AFTableCellSubtotal")) {
                    colorEstilo = RmdColor.decode("#B3C6DB");
                    celdaStyle.setFont(defaultFontBold);
                }                
                
                if (colorEstilo != null) {
                    aplicaColor(celda, celdaStyle, colorEstilo);
                }
                
                HashMap<String, String> mapStyle = getMapaEstilos(style);
                Color color = getColor(mapStyle, "background-color");
                
                if (color != null) {
                    aplicaColor(celda, celdaStyle, color);
                }

                if (align != null) {
                    if (align.contentEquals("start"))
                        celdaStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);
                    else if (align.contentEquals("end"))
                        celdaStyle.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
                    else if (align.contentEquals("center"))
                        celdaStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
                }
                
                esFechaRMD = (columnasFecRMD != null && columnasFecRMD.contains(j));

                for (UIComponent c : col.getChildren()) {
                    if (c instanceof UIXValue) {
                        UIXValue uixValue = (UIXValue)c;
                        if (uixValue.getValue() != null) {
                            if (esFechaRMD) {
                                XSSFDataFormat poiFormat = libro.createDataFormat();
                                celdaStyle.setDataFormat(poiFormat.getFormat("dd/MM/yyyy"));
                                celda.setCellStyle(celdaStyle);
                                fechaRMDTxt = uixValue.getValue().toString();

                                try {
                                    if (!fechaRMDTxt.contentEquals("0")) {
                                        fechaRMD = sdfFormato.parse(fechaRMDTxt);
                                        celda.setCellValue(new java.sql.Date(fechaRMD.getTime()));
                                    }
                                } catch (ParseException e) {
                                    e.getMessage();
                                }
                            } else if (uixValue.getValue() instanceof Number) {
                                celda.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
                                //XSSFDataFormat poiFormat = wb.createDataFormat();
                                //cellStyle.setDataFormat(poiFormat.getFormat("##0.00"));
                                //cell.setCellStyle(cellStyle);
                                celda.setCellValue(((Number)uixValue.getValue()).doubleValue());
                            } else if (uixValue.getValue() instanceof java.sql.Date) {
                                XSSFDataFormat poiFormat = libro.createDataFormat();
                                celdaStyle.setDataFormat(poiFormat.getFormat("dd/MM/yyyy"));
                                celda.setCellStyle(celdaStyle);
                                celda.setCellValue((java.sql.Date)uixValue.getValue());
                            } else
                                celda.setCellValue(uixValue.getValue().toString());
                        }
                    }
                }
            }
        }

        for (int i = 0; i < cols.size(); i++) {
            hoja.autoSizeColumn(i);
        }

        return libro;
    }
}

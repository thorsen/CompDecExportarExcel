package es.ramondin.compdec.exportarexcel.view.beans;

import es.ramondin.compdec.exportarexcel.view.util.ExportarExcelUtil;

import es.ramondin.util.general.RmdMensaje;
import es.ramondin.utilidades.JSFUtils;

import java.io.IOException;
import java.io.OutputStream;

import java.util.ArrayList;

import javax.faces.component.UIViewRoot;
import javax.faces.context.FacesContext;

import oracle.adf.view.rich.component.rich.data.RichTable;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExportarExcelBean {
    public ExportarExcelBean() {
    }

    public void exportarExcel(FacesContext facesContext, OutputStream outputStream) {
        String idTabla = (String)JSFUtils.resolveExpression("#{attrs.idTabla}");
        String nombreHoja = (String)JSFUtils.resolveExpression("#{attrs.nombreHoja}");
        Boolean mostrarColsOcultas = (Boolean)JSFUtils.resolveExpression("#{attrs.mostrarColsOcultas}");
        Integer[] columnasFechaRMD = (Integer[])JSFUtils.resolveExpression("#{attrs.arrayColumnasFechaRMD}");
        
        ArrayList<Integer> columnasFecRMD = null;
        
        int numColumnas = columnasFechaRMD != null ? columnasFechaRMD.length : 0;
        
        if (numColumnas > 0) {
            columnasFecRMD = new ArrayList<Integer>();
        
            for (int i = 0; i < numColumnas; i++)
                columnasFecRMD.add(columnasFechaRMD[i]);
        }

        UIViewRoot vr = facesContext.getViewRoot();
        RichTable tabla = (RichTable)vr.findComponent(idTabla);
        
        XSSFWorkbook wb = ExportarExcelUtil.generarExcel(facesContext, tabla, nombreHoja, mostrarColsOcultas, columnasFecRMD);

        try {
            wb.write(outputStream);
            outputStream.flush();
        } catch (IOException e) {
            RmdMensaje.muestraExcepcion(facesContext, e);
        }
    }
}
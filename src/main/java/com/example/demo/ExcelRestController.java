package com.example.demo;

import com.aspose.cells.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.env.Environment;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Component;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;

@RestController
@RequestMapping("/api")
public class ExcelRestController {

    @Autowired
    private Environment environment;

    @RequestMapping("/createExcel")
    public ResponseEntity<String> createExcel() throws Exception {
        createExcelWithMultiSelectDropDown();
        return new ResponseEntity<>("SUCCESS", HttpStatus.OK );
    }


    public void createExcelWithMultiSelectDropDown() throws Exception {


        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.getWorksheets().get(0);

        int sheetIndex = workbook.getWorksheets().add();
        Worksheet worksheet1 = workbook.getWorksheets().get(sheetIndex);

        Range range = worksheet1.getCells().createRange(0,0,4,1);
        range.setName("MyDropdown");

        range.get(0,0).setValue("Blue");
        range.get(1,0).setValue("Red");
        range.get(2,0).setValue("Green");
        range.get(3,0).setValue("Yellow");

        CellArea area = new CellArea();
        area.StartRow = 0;
        area.StartColumn= 0;
        area.EndRow=4;
        area.EndColumn=0;

        ValidationCollection validation = worksheet.getValidations();

        int index = validation.add(area);
        Validation validation2 = validation.get(index);

        validation2.setType(ValidationType.LIST);
        validation2.setInCellDropDown(true);
        validation2.setFormula1("=MyDropdown");

        int idx = workbook.getVbaProject().getModules().add(worksheet);

        //worksheet.
        VbaModule module2 = workbook.getVbaProject().getModules().get(idx);
        module2.setName("TestModule");
		module2.setCodes("Private Sub Worksheet_Change(ByVal Target As Range)\n" +
				"    Dim xRng As Range\n" +
				"    Dim xValue1 As String\n" +
				"    Dim xValue2 As String\n" +
				"    If Target.Count > 1 Then Exit Sub\n" +
				"    On Error Resume Next\n" +
				"    Set xRng = Cells.SpecialCells(xlCellTypeAllValidation)\n" +
				"    If xRng Is Nothing Then Exit Sub\n" +
				"    Application.EnableEvents = False\n" +
				"    If Not Application.Intersect(Target, xRng) Is Nothing Then\n" +
				"        xValue2 = Target.Value\n" +
				"        Application.Undo\n" +
				"        xValue1 = Target.Value\n" +
				"        Target.Value = xValue2\n" +
				"        If xValue1 <> \"\" Then\n" +
				"            If xValue2 <> \"\" Then\n" +
				"                If xValue1 = xValue2 Or _\n" +
				"                   InStr(1, xValue1, \", \" & xValue2) Or _\n" +
				"                   InStr(1, xValue1, xValue2 & \",\") Then\n" +
				"                    Target.Value = xValue1\n" +
				"                Else\n" +
				"                    Target.Value = xValue1 & \", \" & xValue2\n" +
				"                End If\n" +
				"            End If\n" +
				"        End If\n" +
				"    End If\n" +
				"    Application.EnableEvents = True\n" +
				"End Sub\n");

        workbook.save("Excel.xlsm",SaveFormat.XLSM);

    }
}

package com.example.zunay.survey;

import android.Manifest;
import android.app.Activity;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.os.Environment;
import android.support.v4.app.ActivityCompat;
import android.support.v4.content.ContextCompat;
import android.support.v7.app.AppCompatActivity;
import android.view.View;
import android.widget.CheckBox;
import android.widget.EditText;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class MainActivity extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        try{
            ActivityCompat.requestPermissions(MainActivity.this,
                    new String[]{Manifest.permission.READ_EXTERNAL_STORAGE},
                    1);
            Save("run");
        }catch (IOException e){e.printStackTrace();}

    }

    public void Save(String s) throws IOException {
        File Root = Environment.getExternalStorageDirectory();
        File Dir = new File(Root.getAbsolutePath() + "/Survey");
        if (!Dir.exists()) {
            Dir.mkdir();
            File file = new File(Root.getAbsolutePath() + "/Survey","Survey.xls");

            //New Workbook
            Workbook wb = new HSSFWorkbook();
            Cell c;
            //Cell style for header row
            CellStyle cs = wb.createCellStyle();
            cs.setFillForegroundColor(HSSFColor.LIME.index);
            cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            //New Sheet
            Sheet sheet1;
            sheet1 = wb.createSheet("Survey");
            // Generate column headings
            Row row = sheet1.createRow(0);
            int i = 0;
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input1));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input2));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input3));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input4));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input5));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input6));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input7));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input8));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input9));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input10));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input11));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input12));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input13));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input14));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input15));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input16));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input17));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input18));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input19));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input20));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input21));c.setCellStyle(cs);
            c = row.createCell(i);i++;c.setCellValue(getString(R.string.input22));c.setCellStyle(cs);
            c = row.createCell(i);c.setCellValue(getString(R.string.input23));c.setCellStyle(cs);


            for (int itr = 0; itr < 100; itr++) {
                sheet1.setColumnWidth(itr,(15*500));
            }
            //c = row.createCell(i);c.setCellValue(getString(R.string.input98));c.setCellStyle(cs);
            FileOutputStream fileOutputStream = new FileOutputStream(file,false);
            wb.write(fileOutputStream);
            Toast.makeText(getApplicationContext(),"Saved",Toast.LENGTH_SHORT).show();
        }
        else {
            File file = new File(Root.getAbsolutePath() + "/Survey","Survey.xls");
            if(!file.exists()) {
                Workbook wb = new HSSFWorkbook();
                Cell c;
                CellStyle cs = wb.createCellStyle();
                cs.setFillForegroundColor(HSSFColor.LIME.index);
                cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                Sheet sheet1;
                sheet1 = wb.createSheet("Survey");
                Row row = sheet1.createRow(0);
                int i = 0;
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input1));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input2));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input3));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input4));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input5));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input6));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input7));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input8));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input9));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input10));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input11));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input12));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input13));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input14));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input15));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input16));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input17));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input18));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input19));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input20));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input21));c.setCellStyle(cs);
                c = row.createCell(i);i++;c.setCellValue(getString(R.string.input22));c.setCellStyle(cs);
                c = row.createCell(i);c.setCellValue(getString(R.string.input23));c.setCellStyle(cs);

                for (int itr = 0; itr < 100; itr++) {
                    sheet1.setColumnWidth(itr, (15 * 500));
                }

                FileOutputStream fileOutputStream = new FileOutputStream(file, false);
                wb.write(fileOutputStream);
                Toast.makeText(getApplicationContext(), "Created", Toast.LENGTH_SHORT).show();
            }

        }

    }
    public void Update(View view) throws IOException {
        EditText editText1=(EditText)findViewById(R.id.nationality);
        EditText editText2=(EditText)findViewById(R.id.professionalism);
        EditText editText3=(EditText)findViewById(R.id.workingIn);
        EditText editText4=(EditText)findViewById(R.id.diagnostic_modality);
        EditText editText5=(EditText)findViewById(R.id.lung_mass);
        EditText editText6=(EditText)findViewById(R.id.EGFR_testing);
        EditText editText7=(EditText)findViewById(R.id.Lung_cancer_management);

        CheckBox checkBox1a=(CheckBox)findViewById(R.id.checkBox_nationality_1);
        CheckBox checkBox1b=(CheckBox)findViewById(R.id.checkBox_nationality_2);
        CheckBox checkBox2a=(CheckBox)findViewById(R.id.checkBox_profession_1);
        CheckBox checkBox2b=(CheckBox)findViewById(R.id.checkBox_profession_2);
        CheckBox checkBox2c=(CheckBox)findViewById(R.id.checkBox_profession_3);
        CheckBox checkBox2d=(CheckBox)findViewById(R.id.checkBox_profession_4);
        CheckBox checkBox2e=(CheckBox)findViewById(R.id.checkBox_profession_5);
        CheckBox checkBox2f=(CheckBox)findViewById(R.id.checkBox_profession_6);
        CheckBox checkBox3a=(CheckBox)findViewById(R.id.checkBox_clinical_practice_1);
        CheckBox checkBox3b=(CheckBox)findViewById(R.id.checkBox_clinical_practice_2);
        CheckBox checkBox3c=(CheckBox)findViewById(R.id.checkBox_clinical_practice_3);
        CheckBox checkBox3d=(CheckBox)findViewById(R.id.checkBox_clinical_practice_4);
        CheckBox checkBox4a=(CheckBox)findViewById(R.id.checkBox_work_practice_1);
        CheckBox checkBox4b=(CheckBox)findViewById(R.id.checkBox_work_practice_2);
        CheckBox checkBox4c=(CheckBox)findViewById(R.id.checkBox_work_practice_3);
        CheckBox checkBox4d=(CheckBox)findViewById(R.id.checkBox_work_practice_4);
        CheckBox checkBox5a=(CheckBox)findViewById(R.id.checkBox_persentage_1);
        CheckBox checkBox5b=(CheckBox)findViewById(R.id.checkBox_persentage_2);
        CheckBox checkBox5c=(CheckBox)findViewById(R.id.checkBox_persentage_3);
        CheckBox checkBox6a=(CheckBox)findViewById(R.id.checkBox_lung_cancer_patient_1);
        CheckBox checkBox6b=(CheckBox)findViewById(R.id.checkBox_lung_cancer_patient_2);
        CheckBox checkBox6c=(CheckBox)findViewById(R.id.checkBox_lung_cancer_patient_3);
        CheckBox checkBox7a=(CheckBox)findViewById(R.id.checkBox_existing_lung_cancer_patient_1);
        CheckBox checkBox7b=(CheckBox)findViewById(R.id.checkBox_existing_lung_cancer_patient_2);
        CheckBox checkBox7c=(CheckBox)findViewById(R.id.checkBox_existing_lung_cancer_patient_3);
        CheckBox checkBox8a=(CheckBox)findViewById(R.id.checkBox_common_histology_1);
        CheckBox checkBox8b=(CheckBox)findViewById(R.id.checkBox_common_histology_2);
        CheckBox checkBox8c=(CheckBox)findViewById(R.id.checkBox_common_histology_3);
        CheckBox checkBox8d=(CheckBox)findViewById(R.id.checkBox_common_histology_4);
        CheckBox checkBox8e=(CheckBox)findViewById(R.id.checkBox_common_histology_5);
        CheckBox checkBox9a=(CheckBox)findViewById(R.id.checkBox_metastatic_lung_cancer_patient_1);
        CheckBox checkBox9b=(CheckBox)findViewById(R.id.checkBox_metastatic_lung_cancer_patient_2);
        CheckBox checkBox9c=(CheckBox)findViewById(R.id.checkBox_metastatic_lung_cancer_patient_3);
        CheckBox checkBox10a=(CheckBox)findViewById(R.id.checkBox_diagnostic_modality_1);
        CheckBox checkBox10b=(CheckBox)findViewById(R.id.checkBox_diagnostic_modality_2);
        CheckBox checkBox10c=(CheckBox)findViewById(R.id.checkBox_diagnostic_modality_3);
        CheckBox checkBox10d=(CheckBox)findViewById(R.id.checkBox_diagnostic_modality_4);
        CheckBox checkBox10e=(CheckBox)findViewById(R.id.checkBox_diagnostic_modality_5);
        CheckBox checkBox10f=(CheckBox)findViewById(R.id.checkBox_diagnostic_modality_6);
        CheckBox checkBox10g=(CheckBox)findViewById(R.id.checkBox_diagnostic_modality_7);
        CheckBox checkBox11a=(CheckBox)findViewById(R.id.checkBox_lung_mass_1);
        CheckBox checkBox11b=(CheckBox)findViewById(R.id.checkBox_lung_mass_2);
        CheckBox checkBox11c=(CheckBox)findViewById(R.id.checkBox_lung_mass_3);
        CheckBox checkBox11d=(CheckBox)findViewById(R.id.checkBox_lung_mass_4);
        CheckBox checkBox11e=(CheckBox)findViewById(R.id.checkBox_lung_mass_5);
        CheckBox checkBox11f=(CheckBox)findViewById(R.id.checkBox_lung_mass_6);
        CheckBox checkBox12a=(CheckBox)findViewById(R.id.checkBox_staging_work_up_1);
        CheckBox checkBox12b=(CheckBox)findViewById(R.id.checkBox_staging_work_up_2);
        CheckBox checkBox13a=(CheckBox)findViewById(R.id.checkBox_metastatic_squamous_1);
        CheckBox checkBox13b=(CheckBox)findViewById(R.id.checkBox_metastatic_squamous_2);
        CheckBox checkBox13c=(CheckBox)findViewById(R.id.checkBox_metastatic_squamous_3);
        CheckBox checkBox13d=(CheckBox)findViewById(R.id.checkBox_metastatic_squamous_4);
        CheckBox checkBox13e=(CheckBox)findViewById(R.id.checkBox_metastatic_squamous_5);
        CheckBox checkBox14a=(CheckBox)findViewById(R.id.checkBox_metastatic_adeno_1);
        CheckBox checkBox14b=(CheckBox)findViewById(R.id.checkBox_metastatic_adeno_2);
        CheckBox checkBox14c=(CheckBox)findViewById(R.id.checkBox_metastatic_adeno_3);
        CheckBox checkBox14d=(CheckBox)findViewById(R.id.checkBox_metastatic_adeno_4);
        CheckBox checkBox14e=(CheckBox)findViewById(R.id.checkBox_metastatic_adeno_5);
        CheckBox checkBox15a=(CheckBox)findViewById(R.id.checkBox_cell_carcinoma_1);
        CheckBox checkBox15b=(CheckBox)findViewById(R.id.checkBox_cell_carcinoma_2);
        CheckBox checkBox15c=(CheckBox)findViewById(R.id.checkBox_cell_carcinoma_3);
        CheckBox checkBox15d=(CheckBox)findViewById(R.id.checkBox_cell_carcinoma_4);
        CheckBox checkBox15e=(CheckBox)findViewById(R.id.checkBox_cell_carcinoma_5);
        CheckBox checkBox16a=(CheckBox)findViewById(R.id.checkBox_preferred_TKI_1);
        CheckBox checkBox16b=(CheckBox)findViewById(R.id.checkBox_preferred_TKI_2);
        CheckBox checkBox16c=(CheckBox)findViewById(R.id.checkBox_preferred_TKI_3);
        CheckBox checkBox16d=(CheckBox)findViewById(R.id.checkBox_preferred_TKI_4);
        CheckBox checkBox17a=(CheckBox)findViewById(R.id.checkBox_EGFR_testing_1);
        CheckBox checkBox17b=(CheckBox)findViewById(R.id.checkBox_EGFR_testing_2);
        CheckBox checkBox17c=(CheckBox)findViewById(R.id.checkBox_EGFR_testing_3);
        CheckBox checkBox18a=(CheckBox)findViewById(R.id.checkBox_TKI_blindly_1);
        CheckBox checkBox18b=(CheckBox)findViewById(R.id.checkBox_TKI_blindly_2);
        CheckBox checkBox18c=(CheckBox)findViewById(R.id.checkBox_TKI_blindly_3);
        CheckBox checkBox19a=(CheckBox)findViewById(R.id.checkBox_metastatic_NSCLC_1);
        CheckBox checkBox19b=(CheckBox)findViewById(R.id.checkBox_metastatic_NSCLC_2);
        CheckBox checkBox19c=(CheckBox)findViewById(R.id.checkBox_metastatic_NSCLC_3);
        CheckBox checkBox19d=(CheckBox)findViewById(R.id.checkBox_metastatic_NSCLC_4);
        CheckBox checkBox19e=(CheckBox)findViewById(R.id.checkBox_metastatic_NSCLC_5);
        CheckBox checkBox20a=(CheckBox)findViewById(R.id.checkBox_ECOG_2_1);
        CheckBox checkBox20b=(CheckBox)findViewById(R.id.checkBox_ECOG_2_2);
        CheckBox checkBox21a=(CheckBox)findViewById(R.id.checkBox_immunotherapy_1);
        CheckBox checkBox21b=(CheckBox)findViewById(R.id.checkBox_immunotherapy_2);
        CheckBox checkBox22a=(CheckBox)findViewById(R.id.checkBox_treatment_algorithm_1);
        CheckBox checkBox22b=(CheckBox)findViewById(R.id.checkBox_treatment_algorithm_2);
        CheckBox checkBox22c=(CheckBox)findViewById(R.id.checkBox_treatment_algorithm_3);
        CheckBox checkBox22d=(CheckBox)findViewById(R.id.checkBox_treatment_algorithm_4);
        CheckBox checkBox22e=(CheckBox)findViewById(R.id.checkBox_treatment_algorithm_5);

        String msg1a="";
        String msg1b="";
        String msg2a="";
        String msg2b="";
        String msg2c="";
        String msg2d="";
        String msg2e="";
        String msg2f="";
        String msg3a="";
        String msg3b="";
        String msg3c="";
        String msg4a="";
        String msg4b="";
        String msg4c="";
        String msg4d="";
        String msg5a="";
        String msg5b="";
        String msg5c="";
        String msg6a="";
        String msg6b="";
        String msg6c="";
        String msg7a="";
        String msg7b="";
        String msg7c="";
        String msg8a="";
        String msg8b="";
        String msg8c="";
        String msg8d="";
        String msg8e="";
        String msg9a="";
        String msg9b="";
        String msg9c="";
        String msg10a="";
        String msg10b="";
        String msg10c="";
        String msg10d="";
        String msg10e="";
        String msg10f="";
        String msg10g="";
        String msg11a="";
        String msg11b="";
        String msg11c="";
        String msg11d="";
        String msg11e="";
        String msg11f="";
        String msg12a="";
        String msg12b="";
        String msg13a="";
        String msg13b="";
        String msg13c="";
        String msg13d="";
        String msg13e="";
        String msg14a="";
        String msg14b="";
        String msg14c="";
        String msg14d="";
        String msg14e="";
        String msg15a="";
        String msg15b="";
        String msg15c="";
        String msg15d="";
        String msg15e="";
        String msg16a="";
        String msg16b="";
        String msg16c="";
        String msg16d="";
        String msg17a="";
        String msg17b="";
        String msg17c="";
        String msg18a="";
        String msg18b="";
        String msg18c="";
        String msg19a="";
        String msg19b="";
        String msg19c="";
        String msg19d="";
        String msg19e="";
        String msg20a="";
        String msg20b="";
        String msg21a="";
        String msg21b="";
        String msg22a="";
        String msg22b="";
        String msg22c="";
        String msg22d="";
        String msg22e="";
        String msg23="";


        if(checkBox1a.isChecked()==true)
        {
            msg1a=checkBox1a.getText().toString()+",";
        }
        if(checkBox1b.isChecked()==true)
        {
            msg1b=editText1.getText().toString();
        }
        if(checkBox2a.isChecked()==true)
        {
            msg2a=checkBox2a.getText().toString()+",";
        }
        if(checkBox2b.isChecked()==true)
        {
            msg2b=checkBox2b.getText().toString()+",";
        }
        if(checkBox2c.isChecked()==true)
        {
            msg2c=checkBox2c.getText().toString()+",";
        }
        if(checkBox2d.isChecked()==true)
        {
            msg2d=checkBox2d.getText().toString()+",";
        }
        if(checkBox2e.isChecked()==true)
        {
            msg2e=checkBox2e.getText().toString()+",";
        }
        if(checkBox2f.isChecked()==true)
        {
            msg2f=editText2.getText().toString();
        }
        if(checkBox3a.isChecked()==true)
        {
            msg3a=checkBox3a.getText().toString()+",";
        }
        if(checkBox3b.isChecked()==true)
        {
            msg3b=checkBox3b.getText().toString()+",";
        }
        if(checkBox3c.isChecked()==true)
        {
            msg3c=checkBox3c.getText().toString();
        }
        if(checkBox4a.isChecked()==true)
        {
            msg4a=checkBox4a.getText().toString()+",";
        }
        if(checkBox4b.isChecked()==true)
        {
            msg4b=checkBox4b.getText().toString()+",";
        }
        if(checkBox4c.isChecked()==true)
        {
            msg4c=checkBox4c.getText().toString()+",";
        }
        if(checkBox4d.isChecked()==true)
        {
            msg4d=editText3.getText().toString();
        }
        if(checkBox5a.isChecked()==true)
        {
            msg5a=checkBox5a.getText().toString()+",";
        }
        if(checkBox5b.isChecked()==true)
        {
            msg5b=checkBox5b.getText().toString()+",";
        }
        if(checkBox5c.isChecked()==true)
        {
            msg5c=checkBox5c.getText().toString();
        }
        if(checkBox6a.isChecked()==true)
        {
            msg6a=checkBox6a.getText().toString()+",";
        }
        if(checkBox6b.isChecked()==true)
        {
            msg6b=checkBox6b.getText().toString()+",";
        }
        if(checkBox6c.isChecked()==true)
        {
            msg6c=checkBox6c.getText().toString();
        }
        if(checkBox7a.isChecked()==true)
        {
            msg7a=checkBox7a.getText().toString()+",";
        }
        if(checkBox7b.isChecked()==true)
        {
            msg7b=checkBox7b.getText().toString()+",";
        }
        if(checkBox7c.isChecked()==true)
        {
            msg7c=checkBox7c.getText().toString();
        }
        if(checkBox8a.isChecked()==true)
        {
            msg8a=checkBox8a.getText().toString()+",";
        }
        if(checkBox8b.isChecked()==true)
        {
            msg8b=checkBox8b.getText().toString()+",";
        }
        if(checkBox8c.isChecked()==true)
        {
            msg8c=checkBox8c.getText().toString()+",";
        }
        if(checkBox8d.isChecked()==true)
        {
            msg8d=checkBox8d.getText().toString()+",";
        }
        if(checkBox8e.isChecked()==true)
        {
            msg8e=checkBox8e.getText().toString();
        }
        if(checkBox9a.isChecked()==true)
        {
            msg9a=checkBox9a.getText().toString()+",";
        }
        if(checkBox9b.isChecked()==true)
        {
            msg9b=checkBox9b.getText().toString()+",";
        }
        if(checkBox9c.isChecked()==true)
        {
            msg9c=checkBox9c.getText().toString();
        }
        if(checkBox10a.isChecked()==true)
        {
            msg10a=checkBox10a.getText().toString()+",";
        }
        if(checkBox10b.isChecked()==true)
        {
            msg10b=checkBox10b.getText().toString()+",";
        }
        if(checkBox10c.isChecked()==true)
        {
            msg10c=checkBox10c.getText().toString()+",";
        }
        if(checkBox10d.isChecked()==true)
        {
            msg10d=checkBox10d.getText().toString()+",";
        }
        if(checkBox10e.isChecked()==true)
        {
            msg10e=checkBox10e.getText().toString()+",";
        }
        if(checkBox10f.isChecked()==true)
        {
            msg10f=checkBox10f.getText().toString()+",";
        }
        if(checkBox10g.isChecked()==true)
        {
            msg10g=editText4.getText().toString();
        }
        if(checkBox11a.isChecked()==true)
        {
            msg11a=checkBox11a.getText().toString()+",";
        }
        if(checkBox11b.isChecked()==true)
        {
            msg11b=checkBox11b.getText().toString()+",";
        }
        if(checkBox11c.isChecked()==true)
        {
            msg11c=checkBox11c.getText().toString()+",";
        }
        if(checkBox11d.isChecked()==true)
        {
            msg11d=checkBox11d.getText().toString()+",";
        }
        if(checkBox11e.isChecked()==true)
        {
            msg11e=checkBox11e.getText().toString()+",";
        }
        if(checkBox11f.isChecked()==true)
        {
            msg11f=editText5.getText().toString();
        }
        if(checkBox12a.isChecked()==true)
        {
            msg12a=checkBox12a.getText().toString()+",";
        }
        if(checkBox12b.isChecked()==true)
        {
            msg12b=checkBox12b.getText().toString();
        }
        if(checkBox13a.isChecked()==true)
        {
            msg13a=checkBox13a.getText().toString()+",";
        }
        if(checkBox13b.isChecked()==true)
        {
            msg13b=checkBox13b.getText().toString()+",";
        }
        if(checkBox13c.isChecked()==true)
        {
            msg13c=checkBox13c.getText().toString()+",";
        }
        if(checkBox13d.isChecked()==true)
        {
            msg13d=checkBox13d.getText().toString()+",";
        }
        if(checkBox13e.isChecked()==true)
        {
            msg13e=checkBox13e.getText().toString();
        }
        if(checkBox14a.isChecked()==true)
        {
            msg14a=checkBox14a.getText().toString()+",";
        }
        if(checkBox14b.isChecked()==true)
        {
            msg14b=checkBox14b.getText().toString()+",";
        }
        if(checkBox14c.isChecked()==true)
        {
            msg14c=checkBox14c.getText().toString()+",";
        }
        if(checkBox14d.isChecked()==true)
        {
            msg14d=checkBox14d.getText().toString()+",";
        }
        if(checkBox14e.isChecked()==true)
        {
            msg14e=checkBox14e.getText().toString();
        }
        if(checkBox15a.isChecked()==true)
        {
            msg15a=checkBox15a.getText().toString()+",";
        }
        if(checkBox15b.isChecked()==true)
        {
            msg15b=checkBox15b.getText().toString()+",";
        }
        if(checkBox15c.isChecked()==true)
        {
            msg15c=checkBox15c.getText().toString()+",";
        }
        if(checkBox15d.isChecked()==true)
        {
            msg15d=checkBox15d.getText().toString()+",";
        }
        if(checkBox15e.isChecked()==true)
        {
            msg15e=checkBox15e.getText().toString();
        }
        if(checkBox16a.isChecked()==true)
        {
            msg16a=checkBox16a.getText().toString()+",";
        }
        if(checkBox16b.isChecked()==true)
        {
            msg16b=checkBox16b.getText().toString()+",";
        }
        if(checkBox16c.isChecked()==true)
        {
            msg16c=checkBox16c.getText().toString()+",";
        }
        if(checkBox16d.isChecked()==true)
        {
            msg16d=checkBox16d.getText().toString();
        }
        if(checkBox17a.isChecked()==true)
        {
            msg17a=checkBox17a.getText().toString()+",";
        }
        if(checkBox17b.isChecked()==true)
        {
            msg17b=checkBox17b.getText().toString()+",";
        }
        if(checkBox17c.isChecked()==true)
        {
            msg17c=editText6.getText().toString();
        }
        if(checkBox18a.isChecked()==true)
        {
            msg18a=checkBox18a.getText().toString()+",";
        }
        if(checkBox18b.isChecked()==true)
        {
            msg18b=checkBox18b.getText().toString()+",";
        }
        if(checkBox18c.isChecked()==true)
        {
            msg18c=checkBox18c.getText().toString();
        }
        if(checkBox19a.isChecked()==true)
        {
            msg19a=checkBox19a.getText().toString()+",";
        }
        if(checkBox19b.isChecked()==true)
        {
            msg19b=checkBox19b.getText().toString()+",";
        }
        if(checkBox19c.isChecked()==true)
        {
            msg19c=checkBox19c.getText().toString()+",";
        }
        if(checkBox19d.isChecked()==true)
        {
            msg19d=checkBox19d.getText().toString()+",";
        }
        if(checkBox19e.isChecked()==true)
        {
            msg19e=checkBox19e.getText().toString();
        }
        if(checkBox20a.isChecked()==true)
        {
            msg20a=checkBox20a.getText().toString()+",";
        }
        if(checkBox20b.isChecked()==true)
        {
            msg20b=checkBox20b.getText().toString();
        }
        if(checkBox21a.isChecked()==true)
        {
            msg21a=checkBox21a.getText().toString()+",";
        }
        if(checkBox21b.isChecked()==true)
        {
            msg21b=checkBox21b.getText().toString();
        }
        if(checkBox22a.isChecked()==true)
        {
            msg22a=checkBox22a.getText().toString()+",";
        }
        if(checkBox22b.isChecked()==true)
        {
            msg22b=checkBox22b.getText().toString()+",";
        }
        if(checkBox22c.isChecked()==true)
        {
            msg22c=checkBox22c.getText().toString()+",";
        }
        if(checkBox22d.isChecked()==true)
        {
            msg22d=checkBox22d.getText().toString()+",";
        }
        if(checkBox22e.isChecked()==true)
        {
            msg22e=checkBox22e.getText().toString();
        }
        msg23=editText7.getText().toString();

        File Root = Environment.getExternalStorageDirectory();
        HSSFWorkbook workbook = null;
        FileInputStream file = null;
        try {
            file = new FileInputStream(new File(Root.getAbsolutePath() + "/Survey","Survey.xls"));
            workbook = new HSSFWorkbook(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        int rowCount = workbook.getSheetAt(0).getLastRowNum();
        Row row = workbook.getSheetAt(0).createRow(++rowCount);
        Cell c = null;
        int i = 0;
        c = row.createCell(i);i++;c.setCellValue(msg1a+msg1b);
        c = row.createCell(i);i++;c.setCellValue(msg2a+msg2b+msg2c+msg2d+msg2e+msg2f);
        c = row.createCell(i);i++;c.setCellValue(msg3a+msg3b+msg3c);
        c = row.createCell(i);i++;c.setCellValue(msg4a+msg4b+msg4c+msg4d);
        c = row.createCell(i);i++;c.setCellValue(msg5a+msg5b+msg5c);
        c = row.createCell(i);i++;c.setCellValue(msg6a+msg6b+msg6c);
        c = row.createCell(i);i++;c.setCellValue(msg7a+msg7b+msg7c);
        c = row.createCell(i);i++;c.setCellValue(msg8a+msg8b+msg8c+msg8d+msg8e);
        c = row.createCell(i);i++;c.setCellValue(msg9a+msg9b+msg9c);
        c = row.createCell(i);i++;c.setCellValue(msg10a+msg10b+msg10c+msg10d+msg10e+msg10f+msg10g);
        c = row.createCell(i);i++;c.setCellValue(msg11a+msg11b+msg11c+msg11d+msg11e+msg11f);
        c = row.createCell(i);i++;c.setCellValue(msg12a+msg12b);
        c = row.createCell(i);i++;c.setCellValue(msg13a+msg13b+msg13c+msg13d+msg13e);
        c = row.createCell(i);i++;c.setCellValue(msg14a+msg14b+msg14c+msg14d+msg14e);
        c = row.createCell(i);i++;c.setCellValue(msg15a+msg15b+msg15c+msg15d+msg15e);
        c = row.createCell(i);i++;c.setCellValue(msg16a+msg16b+msg16c+msg16d);
        c = row.createCell(i);i++;c.setCellValue(msg17a+msg17b+msg17c);
        c = row.createCell(i);i++;c.setCellValue(msg18a+msg18b+msg18c);
        c = row.createCell(i);i++;c.setCellValue(msg19a+msg19b+msg19c+msg19d+msg19e);
        c = row.createCell(i);i++;c.setCellValue(msg20a+msg20b);
        c = row.createCell(i);i++;c.setCellValue(msg21a+msg21b);
        c = row.createCell(i);i++;c.setCellValue(msg22a+msg22b+msg22c+msg22d+msg22e);
        c = row.createCell(i);i++;c.setCellValue(msg23);

        try {
            file.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        FileOutputStream outFile =new FileOutputStream(new File(Root.getAbsolutePath() + "/Survey","Survey.xls"));
        workbook.write(outFile);
        if (null != outFile)
            outFile.close();
        Toast.makeText(getApplicationContext(),"Updated",Toast.LENGTH_SHORT).show();
        clearData();



    }
    public void clearData() {
        EditText editText1=(EditText)findViewById(R.id.nationality);
        EditText editText2=(EditText)findViewById(R.id.professionalism);
        EditText editText3=(EditText)findViewById(R.id.workingIn);
        EditText editText4=(EditText)findViewById(R.id.diagnostic_modality);
        EditText editText5=(EditText)findViewById(R.id.lung_mass);
        EditText editText6=(EditText)findViewById(R.id.EGFR_testing);
        EditText editText7=(EditText)findViewById(R.id.Lung_cancer_management);

        CheckBox checkBox1a=(CheckBox)findViewById(R.id.checkBox_nationality_1);
        CheckBox checkBox1b=(CheckBox)findViewById(R.id.checkBox_nationality_2);
        CheckBox checkBox2a=(CheckBox)findViewById(R.id.checkBox_profession_1);
        CheckBox checkBox2b=(CheckBox)findViewById(R.id.checkBox_profession_2);
        CheckBox checkBox2c=(CheckBox)findViewById(R.id.checkBox_profession_3);
        CheckBox checkBox2d=(CheckBox)findViewById(R.id.checkBox_profession_4);
        CheckBox checkBox2e=(CheckBox)findViewById(R.id.checkBox_profession_5);
        CheckBox checkBox2f=(CheckBox)findViewById(R.id.checkBox_profession_6);
        CheckBox checkBox3a=(CheckBox)findViewById(R.id.checkBox_clinical_practice_1);
        CheckBox checkBox3b=(CheckBox)findViewById(R.id.checkBox_clinical_practice_2);
        CheckBox checkBox3c=(CheckBox)findViewById(R.id.checkBox_clinical_practice_3);
        CheckBox checkBox3d=(CheckBox)findViewById(R.id.checkBox_clinical_practice_4);
        CheckBox checkBox4a=(CheckBox)findViewById(R.id.checkBox_work_practice_1);
        CheckBox checkBox4b=(CheckBox)findViewById(R.id.checkBox_work_practice_2);
        CheckBox checkBox4c=(CheckBox)findViewById(R.id.checkBox_work_practice_3);
        CheckBox checkBox4d=(CheckBox)findViewById(R.id.checkBox_work_practice_4);
        CheckBox checkBox5a=(CheckBox)findViewById(R.id.checkBox_persentage_1);
        CheckBox checkBox5b=(CheckBox)findViewById(R.id.checkBox_persentage_2);
        CheckBox checkBox5c=(CheckBox)findViewById(R.id.checkBox_persentage_3);
        CheckBox checkBox6a=(CheckBox)findViewById(R.id.checkBox_lung_cancer_patient_1);
        CheckBox checkBox6b=(CheckBox)findViewById(R.id.checkBox_lung_cancer_patient_2);
        CheckBox checkBox6c=(CheckBox)findViewById(R.id.checkBox_lung_cancer_patient_3);
        CheckBox checkBox7a=(CheckBox)findViewById(R.id.checkBox_existing_lung_cancer_patient_1);
        CheckBox checkBox7b=(CheckBox)findViewById(R.id.checkBox_existing_lung_cancer_patient_2);
        CheckBox checkBox7c=(CheckBox)findViewById(R.id.checkBox_existing_lung_cancer_patient_3);
        CheckBox checkBox8a=(CheckBox)findViewById(R.id.checkBox_common_histology_1);
        CheckBox checkBox8b=(CheckBox)findViewById(R.id.checkBox_common_histology_2);
        CheckBox checkBox8c=(CheckBox)findViewById(R.id.checkBox_common_histology_3);
        CheckBox checkBox8d=(CheckBox)findViewById(R.id.checkBox_common_histology_4);
        CheckBox checkBox8e=(CheckBox)findViewById(R.id.checkBox_common_histology_5);
        CheckBox checkBox9a=(CheckBox)findViewById(R.id.checkBox_metastatic_lung_cancer_patient_1);
        CheckBox checkBox9b=(CheckBox)findViewById(R.id.checkBox_metastatic_lung_cancer_patient_2);
        CheckBox checkBox9c=(CheckBox)findViewById(R.id.checkBox_metastatic_lung_cancer_patient_3);
        CheckBox checkBox10a=(CheckBox)findViewById(R.id.checkBox_diagnostic_modality_1);
        CheckBox checkBox10b=(CheckBox)findViewById(R.id.checkBox_diagnostic_modality_2);
        CheckBox checkBox10c=(CheckBox)findViewById(R.id.checkBox_diagnostic_modality_3);
        CheckBox checkBox10d=(CheckBox)findViewById(R.id.checkBox_diagnostic_modality_4);
        CheckBox checkBox10e=(CheckBox)findViewById(R.id.checkBox_diagnostic_modality_5);
        CheckBox checkBox10f=(CheckBox)findViewById(R.id.checkBox_diagnostic_modality_6);
        CheckBox checkBox10g=(CheckBox)findViewById(R.id.checkBox_diagnostic_modality_7);
        CheckBox checkBox11a=(CheckBox)findViewById(R.id.checkBox_lung_mass_1);
        CheckBox checkBox11b=(CheckBox)findViewById(R.id.checkBox_lung_mass_2);
        CheckBox checkBox11c=(CheckBox)findViewById(R.id.checkBox_lung_mass_3);
        CheckBox checkBox11d=(CheckBox)findViewById(R.id.checkBox_lung_mass_4);
        CheckBox checkBox11e=(CheckBox)findViewById(R.id.checkBox_lung_mass_5);
        CheckBox checkBox11f=(CheckBox)findViewById(R.id.checkBox_lung_mass_6);
        CheckBox checkBox12a=(CheckBox)findViewById(R.id.checkBox_staging_work_up_1);
        CheckBox checkBox12b=(CheckBox)findViewById(R.id.checkBox_staging_work_up_2);
        CheckBox checkBox13a=(CheckBox)findViewById(R.id.checkBox_metastatic_squamous_1);
        CheckBox checkBox13b=(CheckBox)findViewById(R.id.checkBox_metastatic_squamous_2);
        CheckBox checkBox13c=(CheckBox)findViewById(R.id.checkBox_metastatic_squamous_3);
        CheckBox checkBox13d=(CheckBox)findViewById(R.id.checkBox_metastatic_squamous_4);
        CheckBox checkBox13e=(CheckBox)findViewById(R.id.checkBox_metastatic_squamous_5);
        CheckBox checkBox14a=(CheckBox)findViewById(R.id.checkBox_metastatic_adeno_1);
        CheckBox checkBox14b=(CheckBox)findViewById(R.id.checkBox_metastatic_adeno_2);
        CheckBox checkBox14c=(CheckBox)findViewById(R.id.checkBox_metastatic_adeno_3);
        CheckBox checkBox14d=(CheckBox)findViewById(R.id.checkBox_metastatic_adeno_4);
        CheckBox checkBox14e=(CheckBox)findViewById(R.id.checkBox_metastatic_adeno_5);
        CheckBox checkBox15a=(CheckBox)findViewById(R.id.checkBox_cell_carcinoma_1);
        CheckBox checkBox15b=(CheckBox)findViewById(R.id.checkBox_cell_carcinoma_2);
        CheckBox checkBox15c=(CheckBox)findViewById(R.id.checkBox_cell_carcinoma_3);
        CheckBox checkBox15d=(CheckBox)findViewById(R.id.checkBox_cell_carcinoma_4);
        CheckBox checkBox15e=(CheckBox)findViewById(R.id.checkBox_cell_carcinoma_5);
        CheckBox checkBox16a=(CheckBox)findViewById(R.id.checkBox_preferred_TKI_1);
        CheckBox checkBox16b=(CheckBox)findViewById(R.id.checkBox_preferred_TKI_2);
        CheckBox checkBox16c=(CheckBox)findViewById(R.id.checkBox_preferred_TKI_3);
        CheckBox checkBox16d=(CheckBox)findViewById(R.id.checkBox_preferred_TKI_4);
        CheckBox checkBox17a=(CheckBox)findViewById(R.id.checkBox_EGFR_testing_1);
        CheckBox checkBox17b=(CheckBox)findViewById(R.id.checkBox_EGFR_testing_2);
        CheckBox checkBox17c=(CheckBox)findViewById(R.id.checkBox_EGFR_testing_3);
        CheckBox checkBox18a=(CheckBox)findViewById(R.id.checkBox_TKI_blindly_1);
        CheckBox checkBox18b=(CheckBox)findViewById(R.id.checkBox_TKI_blindly_2);
        CheckBox checkBox18c=(CheckBox)findViewById(R.id.checkBox_TKI_blindly_3);
        CheckBox checkBox19a=(CheckBox)findViewById(R.id.checkBox_metastatic_NSCLC_1);
        CheckBox checkBox19b=(CheckBox)findViewById(R.id.checkBox_metastatic_NSCLC_2);
        CheckBox checkBox19c=(CheckBox)findViewById(R.id.checkBox_metastatic_NSCLC_3);
        CheckBox checkBox19d=(CheckBox)findViewById(R.id.checkBox_metastatic_NSCLC_4);
        CheckBox checkBox19e=(CheckBox)findViewById(R.id.checkBox_metastatic_NSCLC_5);
        CheckBox checkBox20a=(CheckBox)findViewById(R.id.checkBox_ECOG_2_1);
        CheckBox checkBox20b=(CheckBox)findViewById(R.id.checkBox_ECOG_2_2);
        CheckBox checkBox21a=(CheckBox)findViewById(R.id.checkBox_immunotherapy_1);
        CheckBox checkBox21b=(CheckBox)findViewById(R.id.checkBox_immunotherapy_2);
        CheckBox checkBox22a=(CheckBox)findViewById(R.id.checkBox_treatment_algorithm_1);
        CheckBox checkBox22b=(CheckBox)findViewById(R.id.checkBox_treatment_algorithm_2);
        CheckBox checkBox22c=(CheckBox)findViewById(R.id.checkBox_treatment_algorithm_3);
        CheckBox checkBox22d=(CheckBox)findViewById(R.id.checkBox_treatment_algorithm_4);
        CheckBox checkBox22e=(CheckBox)findViewById(R.id.checkBox_treatment_algorithm_5);

        editText1.setText("");
        editText2.setText("");
        editText3.setText("");
        editText4.setText("");
        editText5.setText("");
        editText6.setText("");
        editText7.setText("");
        checkBox1a.setChecked(false);
        checkBox1b.setChecked(false);
        checkBox2a.setChecked(false);
        checkBox2b.setChecked(false);
        checkBox2c.setChecked(false);
        checkBox2d.setChecked(false);
        checkBox2e.setChecked(false);
        checkBox2f.setChecked(false);
        checkBox3a.setChecked(false);
        checkBox3b.setChecked(false);
        checkBox3c.setChecked(false);
        checkBox3d.setChecked(false);
        checkBox4a.setChecked(false);
        checkBox4b.setChecked(false);
        checkBox4c.setChecked(false);
        checkBox4d.setChecked(false);
        checkBox5a.setChecked(false);
        checkBox5b.setChecked(false);
        checkBox5c.setChecked(false);
        checkBox6a.setChecked(false);
        checkBox6b.setChecked(false);
        checkBox6c.setChecked(false);
        checkBox7a.setChecked(false);
        checkBox7b.setChecked(false);
        checkBox7c.setChecked(false);
        checkBox8a.setChecked(false);
        checkBox8b.setChecked(false);
        checkBox8c.setChecked(false);
        checkBox8d.setChecked(false);
        checkBox8e.setChecked(false);
        checkBox9a.setChecked(false);
        checkBox9b.setChecked(false);
        checkBox9c.setChecked(false);
        checkBox10a.setChecked(false);
        checkBox10b.setChecked(false);
        checkBox10c.setChecked(false);
        checkBox10d.setChecked(false);
        checkBox10e.setChecked(false);
        checkBox10f.setChecked(false);
        checkBox10g.setChecked(false);
        checkBox11a.setChecked(false);
        checkBox11b.setChecked(false);
        checkBox11c.setChecked(false);
        checkBox11d.setChecked(false);
        checkBox11e.setChecked(false);
        checkBox11f.setChecked(false);
        checkBox12a.setChecked(false);
        checkBox12b.setChecked(false);
        checkBox13a.setChecked(false);
        checkBox13b.setChecked(false);
        checkBox13c.setChecked(false);
        checkBox13d.setChecked(false);
        checkBox13e.setChecked(false);
        checkBox14a.setChecked(false);
        checkBox14b.setChecked(false);
        checkBox14c.setChecked(false);
        checkBox14d.setChecked(false);
        checkBox14e.setChecked(false);
        checkBox15a.setChecked(false);
        checkBox15b.setChecked(false);
        checkBox15c.setChecked(false);
        checkBox15d.setChecked(false);
        checkBox15e.setChecked(false);
        checkBox16a.setChecked(false);
        checkBox16b.setChecked(false);
        checkBox16c.setChecked(false);
        checkBox16d.setChecked(false);
        checkBox17a.setChecked(false);
        checkBox17b.setChecked(false);
        checkBox17c.setChecked(false);
        checkBox18a.setChecked(false);
        checkBox18b.setChecked(false);
        checkBox18c.setChecked(false);
        checkBox19a.setChecked(false);
        checkBox19b.setChecked(false);
        checkBox19c.setChecked(false);
        checkBox19d.setChecked(false);
        checkBox19e.setChecked(false);
        checkBox20a.setChecked(false);
        checkBox20b.setChecked(false);
        checkBox21a.setChecked(false);
        checkBox21b.setChecked(false);
        checkBox22a.setChecked(false);
        checkBox22b.setChecked(false);
        checkBox22c.setChecked(false);
        checkBox22d.setChecked(false);
        checkBox22e.setChecked(false);


    }
}

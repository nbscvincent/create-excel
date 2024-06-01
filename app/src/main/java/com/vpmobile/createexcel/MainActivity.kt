package com.vpmobile.createexcel

import android.content.Context
import android.os.Build
import android.os.Bundle
import android.widget.Toast
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.padding
import androidx.compose.material3.Button
import androidx.compose.material3.Scaffold
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.ui.Modifier
import androidx.compose.ui.unit.dp
import com.vpmobile.createexcel.ui.theme.CreateExcelTheme
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import java.io.IOException
import android.Manifest
import android.content.pm.PackageManager
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.fillMaxHeight
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.ui.Alignment
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContent {
            CreateExcelTheme {
                Scaffold(
                    modifier = Modifier
                        .fillMaxSize()
                ) { it ->
                    Column(
                        modifier = Modifier
                            .padding(it)
                            .fillMaxHeight()
                            .fillMaxWidth(),
                        verticalArrangement = Arrangement.Center,
                        horizontalAlignment = Alignment.CenterHorizontally
                    ) {
                        CreateExcelButton {
                            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R) {
                                createExcelFile(applicationContext)

                            } else {
                                if (ContextCompat.checkSelfPermission(
                                        applicationContext,
                                        Manifest.permission.WRITE_EXTERNAL_STORAGE
                                    ) != PackageManager.PERMISSION_GRANTED
                                ) {
                                    ActivityCompat.requestPermissions(
                                        this@MainActivity,
                                        arrayOf(Manifest.permission.WRITE_EXTERNAL_STORAGE),
                                        1
                                    )
                                } else {
                                    createExcelFile(applicationContext)
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

@Composable
fun CreateExcelButton(onClick: () -> Unit) {
    Button(
        onClick = { onClick() },
        modifier = Modifier.padding(16.dp)
    ) {
        Text(text = "Create Excel File")
    }
}


fun createExcelFile(context: Context) {
    val workbook = XSSFWorkbook()
    val sheet = workbook.createSheet("Sample Sheet")

    // Create row/s
    val header = sheet.createRow(0)
    header.createCell(0).setCellValue("Cell 1 Header")
    header.createCell(1).setCellValue("Cell 2 Header")
    header.createCell(2).setCellValue("Cell 3 Header")

    val rowData = sheet.createRow(1)
    rowData.createCell(0).setCellValue("Cell 1 Data")
    rowData.createCell(1).setCellValue("Cell 2 Data")
    rowData.createCell(2).setCellValue("Cell 3 Data")
    val fileName = "SampleExcelFile.xlsx"

    // Get the directory for the Excel file
    val directory = context.getExternalFilesDir("ExcelFiles")

    // Create directory if it doesn't exist
    if (directory != null && !directory.exists()) {
        if (!directory.mkdirs()) {
            Toast.makeText(context, "Can't create directory", Toast.LENGTH_SHORT).show()
            return
        } else {
            Toast.makeText(context, "Directory created", Toast.LENGTH_SHORT).show()
        }
    }

    // Create the excel file within the directory
    val file = File(directory, fileName)

    var fileOutputStream: FileOutputStream? = null

    try {
        // Write the rows to the excel file
        fileOutputStream = FileOutputStream(file)
        workbook.write(fileOutputStream)
        workbook.close()
        Toast.makeText(context, "Excel file created successfully", Toast.LENGTH_SHORT).show()
    } catch (e: IOException) {
        e.printStackTrace()
        Toast.makeText(context, e.localizedMessage, Toast.LENGTH_SHORT).show()
    } finally {
        fileOutputStream?.close()
    }
}

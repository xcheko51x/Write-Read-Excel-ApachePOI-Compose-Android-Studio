package com.xcheko51x.read_write_excel_compose

import android.annotation.SuppressLint
import android.os.Bundle
import android.os.Environment
import android.util.Log
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.wrapContentSize
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material3.Button
import androidx.compose.material3.Card
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.Scaffold
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.unit.dp
import com.google.accompanist.permissions.ExperimentalPermissionsApi
import com.google.accompanist.permissions.rememberMultiplePermissionsState
import com.xcheko51x.read_write_excel_compose.ui.theme.Read_Write_Excel_ComposeTheme
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException


class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContent {

            Read_Write_Excel_ComposeTheme {
                Scaffold(modifier = Modifier.fillMaxSize()) {
                    RegistroView()
                }
            }
        }
    }
}

@SuppressLint("MutableCollectionMutableState")
@OptIn(ExperimentalPermissionsApi::class)
@Composable
fun RegistroView() {
    var listaUsuarios by remember { mutableStateOf(mutableListOf<Usuario>()) }
    var nombre by remember { mutableStateOf("") }
    var edad by remember { mutableStateOf("") }

    val permissions = rememberMultiplePermissionsState(
        permissions = listOf(
            android.Manifest.permission.WRITE_EXTERNAL_STORAGE,
            android.Manifest.permission.READ_EXTERNAL_STORAGE
        )
    )

    LaunchedEffect(key1 = Unit) {
        permissions.launchMultiplePermissionRequest()
    }

    leerExcel(listaUsuarios)

    Column(
        modifier = Modifier
            .fillMaxSize()
            .padding(
                top = 20.dp,
                start = 8.dp,
                end = 8.dp,
                bottom = 8.dp
            )
    ) {
        OutlinedTextField(
            modifier = Modifier
                .fillMaxWidth()
                .padding(8.dp, 4.dp),
            value = nombre,
            onValueChange =  {
                nombre = it
            },
            label = {
                Text(text = "Nombre")
            }
        )

        OutlinedTextField(
            modifier = Modifier
                .fillMaxWidth()
                .padding(8.dp, 4.dp),
            value = edad,
            onValueChange =  {
                edad = it
            },
            label = {
                Text(text = "Edad")
            }
        )

        Button(
            modifier = Modifier
                .wrapContentSize()
                .align(Alignment.CenterHorizontally),
            onClick = {
                val usuario = Usuario(
                    nombre = nombre,
                    edad = edad
                )

                listaUsuarios.add(usuario)

                nombre = ""
                edad = ""
            }
        ) {
            Text(text = "Agregar Registro")
        }

        Button(
            modifier = Modifier
                .wrapContentSize()
                .align(Alignment.CenterHorizontally),
            onClick = {
                crearExcel(listaUsuarios)
            }
        ) {
            Text(text = "Crear Excel")
        }

        LazyColumn(
            modifier = Modifier
                .padding(8.dp)
        ) {
            items(listaUsuarios) {
                Card(
                    modifier = Modifier
                        .fillMaxSize()
                        .padding(8.dp),
                    shape = RoundedCornerShape(8.dp)
                ) {
                    Column(
                        modifier = Modifier
                            .fillMaxSize()
                            .padding(8.dp),
                    ) {
                        Text(text = "Nombre: ${it.nombre}")
                        Text(text = "Edad: ${it.edad}")
                    }
                }
            }
        }
    }
}

fun crearExcel(listaRegistros: MutableList<Usuario>) {

    val path = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS)

    val fileName = "registros.xlsx"

    // Crear un nuevo libro de trabajo Excel en formato .xlsx
    val workbook = XSSFWorkbook()

    // Crear una hoja de trabajo (worksheet)
    val sheet: Sheet = workbook.createSheet("Hoja 1")

    // Crear una fila en la hoja
    val headerRow = sheet.createRow(0)

    // Crear celdas en la fila
    var cell = headerRow.createCell(0)
    cell.setCellValue("Nombre")

    cell = headerRow.createCell(1)
    cell.setCellValue("Edad")

    for (index in listaRegistros.indices) {
        val row = sheet.createRow(index + 1)
        row.createCell(0).setCellValue(listaRegistros[index].nombre)
        row.createCell(1).setCellValue(listaRegistros[index].edad)
    }

    // Guardar el libro de trabajo (workbook) en almacenamiento externo
    try {
        val fileOutputStream = FileOutputStream(
            File(path, fileName)
        )
        workbook.write(fileOutputStream)
        fileOutputStream.close()
        workbook.close()
    } catch (e: IOException) {
        e.printStackTrace()
    }

}

fun leerExcel(listaRegistros: MutableList<Usuario>) {
    val fileName = "registros.xlsx"

    val path = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS).absolutePath+"/"+fileName

    Log.d("EXCEL","EXCEL LEER")

    val lista = arrayListOf<String>()

    try {
        val fileInputStream = FileInputStream(path)
        val workbook = WorkbookFactory.create(fileInputStream)
        val sheet: Sheet = workbook.getSheetAt(0)

        val rows = sheet.iterator()
        while (rows.hasNext()) {
            val currentRow = rows.next()

            // Iterar sobre celdas de la fila actual
            val cellsInRow = currentRow.iterator()
            while (cellsInRow.hasNext()) {
                val currentCell = cellsInRow.next()

                // Obtener valor de la celda como String
                val cellValue: String = when (currentCell.cellType) {
                    CellType.STRING -> currentCell.stringCellValue
                    CellType.NUMERIC -> currentCell.numericCellValue.toString()
                    CellType.BOOLEAN -> currentCell.booleanCellValue.toString()
                    else -> ""
                }

                lista.add(cellValue)

                //Log.d("ExcelReader", "Valor de celda: $cellValue")
            }
        }

        for (i in 2 until lista.size step 2) {
            listaRegistros.add(
                Usuario(
                    lista[i],
                    lista[i+1]
                )
            )
        }

        workbook.close()
        fileInputStream.close()
    } catch (e: IOException) {
        e.printStackTrace()
    }
}

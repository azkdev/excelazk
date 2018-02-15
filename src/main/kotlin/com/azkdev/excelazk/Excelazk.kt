package com.azkdev.excelazk

import com.azkdev.excelazk.extensions.get
import com.azkdev.excelazk.extensions.set
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.IOException
import java.nio.file.Files
import java.nio.file.Paths

fun main(args: Array<String>) {

    // Opening Excel file.
    Excelazk.open("raport.xlsx").use { workbook ->

        val sheet: Sheet = workbook.getSheetAt(0) // Get Sheet 0.

        // Reading some data from file.
        println(sheet[0, 0]) // Read Column 0, Row 0.
        println(sheet[1, 2]) // Read Column 1, Row 2.

        // Writing some data and save a new copy of file.
        sheet[3, 0] = "New value from Kotlin" // Write data to Cell 3, Row 0.

        Excelazk.write(workbook, "raport_2.xlsx") // Write new file with given workbook and name.
    }
}

class Excelazk {
    companion object {
        fun open(filename: String): Workbook {
            return WorkbookFactory.create(Paths.get(filename).toFile())
        }

        fun write(workbook: Workbook, fileName: String) {
            val outPath = Paths.get(fileName)
            try {
                Files.newOutputStream(outPath).use {
                    workbook.write(it)
                }
            } catch (e: IOException) {
                e.printStackTrace()
            }
        }
    }
}
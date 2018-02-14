package com.azkdev.excelazk

import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.FileInputStream
import java.io.IOException
import java.nio.file.Files
import java.nio.file.Paths

fun main(args: Array<String>) {
    println("Excelazk initial commit!")
}

class Excelazk {
    companion object {
        fun open(filename: String): Workbook {
            return WorkbookFactory.create(FileInputStream(Paths.get(filename).toFile()))
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
package com.azkdev.excelazk.extensions

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Row

operator fun Sheet.get(n: Int): Row {
    return this.getRow(n) ?: this.createRow(n)
}

operator fun Sheet.get(x: Int, y: Int): Cell {
    val row: Row = this[y]
    return row[x]
}

operator fun Row.get(n: Int): Cell {
    return this.getCell(n) ?: this.createCell(n)
}

operator fun Sheet.set(x: Int, y: Int, value: Any) {
    val cell: Cell = this[x, y]
    when (value) {
        is String -> cell.setCellValue(value)
        is Int -> cell.setCellValue(value.toDouble())
        is Double -> cell.setCellValue(value)
        else -> throw IllegalArgumentException("Illegal argument!")
    }
}